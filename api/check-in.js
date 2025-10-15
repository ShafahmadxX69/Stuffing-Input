import { GoogleSpreadsheet } from 'google-spreadsheet';

export default async function handler(req, res) {
  if (req.method !== 'POST')
    return res.status(405).json({ error: 'Method Not Allowed' });

  try {
    const { invoice, items } = req.body || {};
    const invoiceTitle = (invoice || '').trim();
    const itemList = Array.isArray(items) ? items : [];

    if (!invoiceTitle)
      return res.status(400).json({ error: 'Invoice title is required' });

    // ====== ENV VARIABLES ======
    const SERVICE_ACCOUNT_EMAIL = process.env.SA_CLIENT_EMAIL;
    const PRIVATE_KEY =
      process.env.SA_PRIVATE_KEY && process.env.SA_PRIVATE_KEY.replace(/\\n/g, '\n');

    // sheet dari link kamu langsung:
    const SHEET_ID = '1XoV7020NTZk1kzqn3F2ks3gOVFJ5arr5NVgUdewWPNQ';
    const SHEET_NAME = 'IN';

    if (!SERVICE_ACCOUNT_EMAIL || !PRIVATE_KEY) {
      return res.status(500).json({ error: 'Missing Google service account credentials' });
    }

    // ====== CONNECT TO SHEET ======
    const doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth({
      client_email: SERVICE_ACCOUNT_EMAIL,
      private_key: PRIVATE_KEY,
    });
    await doc.loadInfo();

    const sheet = doc.sheetsByTitle[SHEET_NAME];
    if (!sheet) return res.status(500).json({ error: `Sheet "${SHEET_NAME}" not found` });

    await sheet.loadHeaderRow();
    const headers = sheet.headerValues.map(h => String(h).trim());
    const rows = await sheet.getRows({ limit: 10000 });

    // Cari kolom yang sesuai dengan judul invoice
    const invoiceColIndex = headers.findIndex(
      (h) => h.toLowerCase() === invoiceTitle.toLowerCase()
    );

    if (invoiceColIndex === -1) {
      return res.json({
        invoiceFound: false,
        summary: `Invoice header "${invoiceTitle}" tidak ditemukan di sheet.`,
        items: [],
      });
    }

    // ====== PROSES ITEM ======
    const results = [];
    for (const it of itemList) {
      const matchRow = rows.find((r) => {
        const uliCell = (
          r['ULI PO'] ||
          r['ULI_PO'] ||
          r['ULIPO'] ||
          r[headers[0]] ||
          ''
        ).toString().trim();

        const materialCell = (
          r['Material No'] ||
          r['MaterialNo'] ||
          r['Material'] ||
          r[headers[2]] ||
          ''
        ).toString().trim();

        const brandCell = (
          r['Brand'] ||
          r['brand'] ||
          r[headers[3]] ||
          ''
        ).toString().trim();

        const colorCell = (
          r['Code Color'] ||
          r['Color Code'] ||
          r['code color'] ||
          r['CodeColor'] ||
          r['Color'] ||
          ''
        ).toString().trim();

        const matchUli = it.uliPO
          ? uliCell.toLowerCase() === it.uliPO.toLowerCase()
          : true;
        const matchMat = it.materialNo
          ? materialCell.toLowerCase() === it.materialNo.toLowerCase()
          : false;
        const matchBrand = it.brand
          ? brandCell.toLowerCase() === it.brand.toLowerCase()
          : true;
        const matchColor = it.codeColor
          ? colorCell.toLowerCase().includes(it.codeColor.toLowerCase())
          : true;

        return matchUli && matchMat && matchBrand && matchColor;
      });

      if (!matchRow) {
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'missing',
          message: 'Tidak ditemukan baris yang cocok di sheet IN',
        });
        continue;
      }

      const invoiceHeader = headers[invoiceColIndex];
      const invoiceQtyRaw = matchRow[invoiceHeader];
      const invoiceQty = invoiceQtyRaw ? Number(invoiceQtyRaw) : 0;
      const expectedQty = Number(it.qty || 0);

      if (invoiceQty === expectedQty) {
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'ok',
          message: `Qty cocok (${invoiceQty})`,
        });
      } else {
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'mismatch',
          message: `Qty di invoice: ${invoiceQty} — expected: ${expectedQty}`,
        });
      }
    }

    // ====== RINGKASAN ======
    const okCount = results.filter((r) => r.status === 'ok').length;
    const missingCount = results.filter((r) => r.status === 'missing').length;
    const misCount = results.filter((r) => r.status === 'mismatch').length;

    return res.json({
      invoiceFound: true,
      summary: `${okCount} cocok, ${misCount} beda qty, ${missingCount} tidak ditemukan`,
      items: results,
    });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message || String(err) });
  }
}

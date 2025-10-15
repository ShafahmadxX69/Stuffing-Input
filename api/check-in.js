// /api/check-in.js
const { GoogleSpreadsheet } = require('google-spreadsheet');

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');
  try{
    const body = req.body;
    const invoiceTitle = (body.invoice || '').trim();
    const items = Array.isArray(body.items) ? body.items : [];

    if(!invoiceTitle) return res.status(400).json({ error: 'Invoice title is required' });

    // Load credentials from env vars (set in Vercel)
    const SERVICE_ACCOUNT_EMAIL = process.env.SA_CLIENT_EMAIL;
    const PRIVATE_KEY = process.env.SA_PRIVATE_KEY && process.env.SA_PRIVATE_KEY.replace(/\\n/g, '\n');
    const SHEET_ID = process.env.SHEET_ID; // Google Spreadsheet ID
    const SHEET_NAME = process.env.SHEET_NAME || 'IN';

    if(!SERVICE_ACCOUNT_EMAIL || !PRIVATE_KEY || !SHEET_ID){
      return res.status(500).json({ error: 'Service account env vars missing' });
    }

    const doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth({
      client_email: SERVICE_ACCOUNT_EMAIL,
      private_key: PRIVATE_KEY
    });
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[SHEET_NAME];
    if(!sheet) return res.status(500).json({ error: `Sheet named "${SHEET_NAME}" not found` });

    // load header row and all rows
    await sheet.loadHeaderRow(); // loads sheet.headerValues
    const headers = sheet.headerValues; // array
    // Find column index for invoiceTitle in headers
    const invoiceColIndex = headers.findIndex(h => String(h).trim() === invoiceTitle);
    if(invoiceColIndex === -1){
      return res.json({ invoiceFound:false, summary: 'Invoice header not found in sheet', items: [] });
    }

    // load all rows (slow for huge sheets, but fine for moderate sizes)
    const rows = await sheet.getRows({ limit: 10000 }); // each row is object keyed by header
    const results = [];
    for(const it of items){
      // find matching row by ULI PO (col A header?), Material No (col C), Brand (col D), codeColor (col H)
      // We'll do case-insensitive match on strings.
      const matchRow = rows.find(r => {
        const uli = (r[ headers[0] ] || '') + ''; // trust header[0] is ULI PO but we also allow other mapping
        // To be robust, compare common columns explicitly if available:
        const uliCell = (r['ULI PO'] || r['ULI_PO'] || r['ULIPO'] || r[headers[0]] || '').toString().trim();
        const materialCell = (r['Material No'] || r['MaterialNo'] || r['Material'] || r[headers[2]] || r['C'] || '').toString().trim();
        const brandCell = (r['Brand'] || r['brand'] || r[headers[3]] || '').toString().trim();
        const colorCell = (r['Code Color'] || r['Color Code'] || r['code color'] || r['H'] || r['CodeColor'] || r['Color'] || '').toString().trim();

        const matchUli = it.uliPO ? uliCell.toLowerCase() === it.uliPO.toLowerCase() : true;
        const matchMat = it.materialNo ? materialCell.toLowerCase() === it.materialNo.toLowerCase() : false;
        const matchBrand = it.brand ? brandCell.toLowerCase() === it.brand.toLowerCase() : true;
        const matchColor = it.codeColor ? colorCell.toLowerCase().includes(it.codeColor.toLowerCase()) : true;

        return matchUli && matchMat && matchBrand && matchColor;
      });

      if(!matchRow){
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'missing',
          message: 'Tidak ditemukan baris yang cocok di sheet IN'
        });
        continue;
      }

      // get invoice qty value on this matched row
      const invoiceHeader = headers[invoiceColIndex];
      const invoiceQtyRaw = matchRow[invoiceHeader];
      const invoiceQty = invoiceQtyRaw ? Number(invoiceQtyRaw) : 0;
      const expectedQty = Number(it.qty || 0);

      if(invoiceQty === expectedQty){
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'ok',
          message: `Qty cocok (${invoiceQty})`
        });
      } else {
        results.push({
          rowIndex: it.rowIndex,
          materialNo: it.materialNo,
          status: 'mismatch',
          message: `Qty di invoice: ${invoiceQty} — expected: ${expectedQty}`
        });
      }
    }

    const okCount = results.filter(r=>r.status==='ok').length;
    const missingCount = results.filter(r=>r.status==='missing').length;
    const misCount = results.filter(r=>r.status==='mismatch').length;

    return res.json({
      invoiceFound: true,
      summary: `${okCount} ok, ${misCount} mismatch, ${missingCount} missing`,
      items: results
    });

  }catch(err){
    console.error(err);
    return res.status(500).json({ error: err.message || String(err) });
  }
};

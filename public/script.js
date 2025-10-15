// script.js
const selectBtn = document.getElementById('selectBtn');
const fileInput = document.getElementById('fileInput');
const dropzone = document.getElementById('dropzone');
const summary = document.getElementById('summary');
const results = document.getElementById('results');

selectBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e)=> handleFiles(e.target.files));

;['dragenter','dragover'].forEach(ev=>{
  dropzone.addEventListener(ev, (e)=>{ e.preventDefault(); dropzone.classList.add('hover') });
});
;['dragleave','drop'].forEach(ev=>{
  dropzone.addEventListener(ev, (e)=>{ e.preventDefault(); dropzone.classList.remove('hover') });
});

dropzone.addEventListener('drop', (e)=>{
  const dt = e.dataTransfer;
  if(!dt) return;
  handleFiles(dt.files);
});

async function handleFiles(fileList){
  if(!fileList || fileList.length===0) return;
  summary.innerHTML = `Memproses ${fileList.length} file...`;
  results.innerHTML = '';
  for(const file of fileList){
    if(!file.name.match(/\.(xlsx|xls)$/i)){
      appendResultRow(`${file.name} bukan file Excel`, 'fail');
      continue;
    }
    try{
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {type:'array'});
      // gunakan sheet pertama
      const sheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[sheetName];
      // baca sebagai array-of-arrays
      const raw = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:false});
      const parsed = parseStuffingSheet(raw);
      displayParsedSummary(file.name, parsed);
      // kirim ke server untuk validasi dengan sheet IN
      const resp = await postToServer(parsed);
      displayServerResponse(file.name, resp);
    }catch(err){
      console.error(err);
      appendResultRow(`Gagal memproses ${file.name}: ${err.message}`, 'fail');
    }
  }
  summary.innerHTML = `Selesai memproses file.`;
}

function parseStuffingSheet(raw){
  // raw is array-of-arrays, 0-indexed: row1 => raw[0]
  const getCell = (r,c) => (raw[r] && raw[r][c] !== undefined) ? String(raw[r][c]).trim() : '';
  const invoice = getCell(2,2); // Row 3, Col C
  const brandTo = getCell(4,2); // Row 5, Col C
  const container = getCell(7,2); // Row 8, Col C

  // items start from row 15 => index 14
  const items = [];
  for(let r = 14; r < raw.length; r++){
    const row = raw[r];
    if(!row) continue;
    const materialNo = row[1] ? String(row[1]).trim() : ''; // B
    if(!materialNo) continue; // skip empty rows
    const modelSize = row[2] ? String(row[2]).trim() : ''; // C
    const qty = row[4] ? Number(row[4]) : 0; // E
    const colorFull = row[5] ? String(row[5]).trim() : ''; // F
    const customerPO = row[6] ? String(row[6]).trim() : ''; // G
    const uliPO = row[7] ? String(row[7]).trim() : ''; // H
    const brand = row[8] ? String(row[8]).trim() : ''; // I

    // extract code color: assume pattern like "Y03#啞光藍 Navy Blue 100665MNVYV1"
    let codeColor = '';
    const match = colorFull.match(/^([A-Za-z0-9#\-]+)/);
    if(match) codeColor = match[1];

    items.push({
      materialNo, modelSize, qty, colorFull, codeColor, customerPO, uliPO, brand, rowIndex: r+1
    });
  }

  return { invoice, brandTo, container, items };
}

function displayParsedSummary(filename, parsed){
  const div = document.createElement('div');
  div.className = 'card';
  div.innerHTML = `<strong>${filename}</strong><br>
    Invoice: ${parsed.invoice || '<i>kosong</i>'} — Brand/TO: ${parsed.brandTo || '<i>kosong</i>'} — Container: ${parsed.container || '<i>kosong</i>'}
    <br>Items found: ${parsed.items.length}`;
  results.appendChild(div);
}

async function postToServer(parsed){
  const res = await fetch('/api/check-in', {
    method:'POST',
    headers: { 'Content-Type':'application/json' },
    body: JSON.stringify(parsed)
  });
  if(!res.ok){
    const text = await res.text();
    throw new Error(text || 'Server error');
  }
  return res.json();
}

function displayServerResponse(filename, resp){
  const div = document.createElement('div');
  div.className = 'card';
  const lines = [];
  lines.push(`<strong>${filename}</strong> — invoice: <em>${resp.invoiceFound? 'Ditemukan' : 'Tidak ditemukan'}</em>`);
  if(!resp.invoiceFound){
    lines.push(`<div class="fail">Invoice tidak ditemukan pada sheet IN</div>`);
  } else {
    lines.push(`<div>Summary: ${resp.summary}</div>`);
    if(resp.items && resp.items.length){
      lines.push('<ul>');
      for(const it of resp.items){
        let cls='success';
        if(it.status === 'ok') cls='success';
        if(it.status === 'missing') cls='warn';
        if(it.status === 'mismatch') cls='fail';
        lines.push(`<li class="${cls}">Row ${it.rowIndex}: ${it.materialNo} — ${it.status} ${it.message? ' — '+it.message : ''}</li>`);
      }
      lines.push('</ul>');
    }
  }
  div.innerHTML = lines.join('\n');
  results.appendChild(div);
}

function appendResultRow(text, cls='') {
  const d = document.createElement('div');
  d.className = 'card';
  if(cls) d.classList.add(cls);
  d.textContent = text;
  results.appendChild(d);
}

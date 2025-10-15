const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('fileInput');
const output = document.getElementById('output');

// Click upload
dropZone.addEventListener('click', () => fileInput.click());

// Drag & drop
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.style.backgroundColor = '#e8f0fe';
});

dropZone.addEventListener('dragleave', () => {
  dropZone.style.backgroundColor = 'white';
});

dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.style.backgroundColor = 'white';
  handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', e => handleFiles(e.target.files));

function handleFiles(files) {
  output.innerHTML = '';
  Array.from(files).forEach(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        renderSheet(file.name, sheetName, json);
      });
    };
    reader.readAsArrayBuffer(file);
  });
}

function renderSheet(filename, sheetName, rows) {
  const container = document.createElement('div');
  container.classList.add('sheet-block');
  container.innerHTML = `
    <h3>${filename} → ${sheetName}</h3>
    <pre>${JSON.stringify(rows.slice(0, 15), null, 2)}</pre>
  `;
  output.appendChild(container);
}

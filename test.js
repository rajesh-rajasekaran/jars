<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Folder-Based Multi-Excel Merger</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root{
    --c1:#e3f2fd; --c2:#f3e5f5; --c3:#e8f5e9; --c4:#fff3e0;
    --c5:#fce4ec; --c6:#e0f7fa; --c7:#f1f8e9; --c8:#fff8e1; --c9:#e1f5fe;
  }
  body{font-family:Arial,Helvetica,sans-serif;background:#f0f4f8;color:#2c3e50;margin:0}
  .container{max-width:1250px;margin:30px auto;background:#fff;border-radius:12px;
             box-shadow:0 8px 25px rgba(0,0,0,.1);overflow:hidden}
  header{background:#2c3e50;color:#fff;padding:20px;text-align:center}
  h1{margin:0;font-size:1.8rem}
  .subtitle{margin:8px 0 0;opacity:.9}
  .controls{padding:20px;background:#f8f9fa;border-bottom:1px solid #dee2e6;
             display:flex;gap:15px;align-items:center;flex-wrap:wrap}
  button{padding:10px 18px;border:none;border-radius:6px;font-weight:bold;
         cursor:pointer;transition:.2s}
  .btn-pick{background:#27ae60;color:#fff}
  .btn-pick:hover{background:#1e8449}
  .btn-clear{background:#e74c3c;color:#fff}
  .btn-clear:hover{background:#c0392b}
  .search{padding:15px 20px;background:#f8f9fa;border-bottom:1px solid #dee2e6}
  .search input{width:100%;padding:12px;border:1px solid #ced4da;border-radius:8px;font-size:1rem}
  .summary{padding:0 20px 15px;font-weight:bold;color:#2980b9;font-size:1.1rem}
  .legend{padding:15px 20px;background:#f1f3f5;border-bottom:1px solid #dee2e6;
           display:flex;flex-wrap:wrap;gap:12px;font-size:.9rem}
  .legend-item{display:flex;align-items:center;gap:6px}
  .legend-color{width:16px;height:16px;border-radius:4px;border:1px solid #ccc}
  .table-wrap{padding:20px;max-height:650px;overflow:auto}
  table{width:100%;border-collapse:collapse;font-size:.95rem}
  th,td{border:1px solid #dee2e6;padding:10px;text-align:left}
  th{background:#2c3e50;color:#fff;position:sticky;top:0;z-index:10}
  tr:nth-child(even) td{background:#f9f9fb}
  .no-data{text-align:center;color:#7f8c8d;font-style:italic;padding:40px;font-size:1.1rem}
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Folder-Based Excel Merger</h1>
    <p class="subtitle">Pick a folder → every *.xlsx/*.xls inside is merged with colour coding</p>
  </header>

  <div class="controls">
    <button id="pickFolder" class="btn-pick">Pick Folder</button>
    <button id="clearAll" class="btn-clear">Clear All</button>
  </div>

  <div class="search"><input id="searchBox" placeholder="Search across all files…"></div>

  <div id="summary" class="summary"></div>
  <div id="legend" class="legend"></div>

  <div class="table-wrap"><div id="output"></div></div>
</div>

<script>
/* ---------- core data ---------- */
let merged = [];               // [{row:[...], file, sheet}]
let filtered = [];
const fileColor = {};
const colors = ['var(--c1)','var(--c2)','var(--c3)','var(--c4)','var(--c5)',
                'var(--c6)','var(--c7)','var(--c8)','var(--c9)'];

/* ---------- UI elements ---------- */
const pickBtn   = document.getElementById('pickFolder');
const clearBtn  = document.getElementById('clearAll');
const searchIn  = document.getElementById('searchBox');
const outDiv    = document.getElementById('output');
const sumDiv    = document.getElementById('summary');
const legDiv    = document.getElementById('legend');

/* ---------- pick folder & read files ---------- */
pickBtn.onclick = async () => {
  try {
    const dirHandle = await window.showDirectoryPicker();
    const promises = [];

    for await (const entry of dirHandle.values()) {
      if (entry.kind !== 'file') continue;
      const name = entry.name.toLowerCase();
      if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) continue;

      promises.push((async () => {
        const file = await entry.getFile();
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, {type:'array'});
        return {fileName: file.name, workbook: wb};
      })());
    }

    const results = await Promise.all(promises);
    if (!results.length) { alert('No Excel files found in the selected folder.'); return; }

    merged = [];   // reset
    results.forEach(({fileName, workbook}) => {
      if (!fileColor[fileName]) {
        fileColor[fileName] = colors[Object.keys(fileColor).length % colors.length];
      }
      workbook.SheetNames.forEach(sn => {
        const ws = workbook.Sheets[sn];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1});
        rows.forEach((r, i) => {
          if (i===0) return;               // skip header row
          merged.push({row: r, file:fileName, sheet:sn});
        });
      });
    });

    filtered = [...merged];
    renderAll();
  } catch (e) {
    if (e.name !== 'AbortError') alert('Error: ' + e.message);
  }
};

/* ---------- clear ---------- */
clearBtn.onclick = () => {
  if (!merged.length || confirm('Remove all data?')) {
    merged = []; filtered = []; fileColor = {};
    renderAll();
  }
};

/* ---------- search ---------- */
searchIn.oninput = () => {
  const q = searchIn.value.trim().toLowerCase();
  filtered = q ? merged.filter(it =>
    it.row.some(c => c!=null && c.toString().toLowerCase().includes(q)) ||
    it.file.toLowerCase().includes(q) ||
    it.sheet.toLowerCase().includes(q)
  ) : [...merged];
  renderTable(); updateSummary();
};

/* ---------- rendering ---------- */
function renderAll(){ renderLegend(); renderTable(); updateSummary(); }

function renderLegend(){
  legDiv.innerHTML = Object.keys(fileColor).length
    ? Object.entries(fileColor).map(([f,c])=>`
        <div class="legend-item">
          <div class="legend-color" style="background:${c}"></div>
          <span>${esc(f)}</span>
        </div>`).join('')
    : '<i>No files loaded</i>';
}

function renderTable(){
  if (!filtered.length){
    outDiv.innerHTML = '<p class="no-data">Pick a folder to load Excel files.</p>';
    return;
  }

  /* master header – Source File, Sheet, then every column that ever appears */
  const hdr = new Set(['Source File','Sheet']);
  filtered.forEach(it=>it.row.forEach((_,i)=>hdr.add(`Col ${i+1}`)));
  const header = Array.from(hdr);

  let html = `<table><thead><tr>${header.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>`;

  filtered.forEach(it=>{
    const bg = fileColor[it.file] || '#fff';
    html += `<tr style="background:${bg};">`;
    html += `<td><strong>${esc(it.file)}</strong></td><td>${esc(it.sheet)}</td>`;
    it.row.forEach(cell=>html+=`<td>${esc(cell)}</td>`);
    const missing = header.length - 2 - it.row.length;
    for(let i=0;i<missing;i++) html += '<td></td>';
    html += '</tr>';
  });
  html += '</tbody></table>';
  outDiv.innerHTML = html;
}

function updateSummary(){
  sumDiv.textContent = `Showing ${filtered.length} of ${merged.length} total rows`;
}

function esc(t){               // simple HTML escape
  if (t==null) return '';
  const d=document.createElement('div');
  d.textContent=t;
  return d.innerHTML;
}
</script>
</body>
</html>
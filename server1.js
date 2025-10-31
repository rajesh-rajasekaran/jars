#!/usr/bin/env node
/*  Excel Grid Viewer – Zero Dependencies  */
/*  Features: Upload (browser/directory), Search, Pagination, Export, Summaries  */

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3000;
const UPLOAD_DIR = path.resolve(__dirname, 'uploads');
const TEMP_DIR = path.resolve(__dirname, 'temp');
[UPLOAD_DIR, TEMP_DIR].forEach(d => fs.existsSync(d) || fs.mkdirSync(d));

// ====================== TINY XLSX PARSER ======================
function parseXLSX(buffer) {
  const zip = unzip(buffer);
  const wbXml = new TextDecoder().decode(zip.get('xl/workbook.xml') || new Uint8Array());
  const sheetNames = [...wbXml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*sheetId="(\d+)"/g)]
    .map(m => m[1]);
  if (!sheetNames.length) throw new Error('No sheets');

  const relsXml = new TextDecoder().decode(zip.get('xl/_rels/workbook.xml.rels') || new Uint8Array());
  const sheetPathMatch = relsXml.match(/Target="([^"]*worksheets\/sheet1\.xml)"/);
  if (!sheetPathMatch) throw new Error('No sheet1');
  const sheetXml = new TextDecoder().decode(zip.get('xl/' + sheetPathMatch[1]) || new Uint8Array());

  const rows = parseSheetXml(sheetXml);
  return { header: rows.header, data: rows.data };
}

function unzip(buf) {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let pos = buf.byteLength - 22;
  while (pos > 0 && view.getUint32(pos, true) !== 0x06054b50) pos--;
  if (pos <= 0) throw new Error('Not a zip file');
  const cdOffset = view.getUint32(pos + 16, true);
  const entries = view.getUint16(pos + 10, true);
  pos = cdOffset;
  const map = new Map();

  for (let i = 0; i < entries; i++) {
    if (view.getUint32(pos, true) !== 0x02014b50) throw new Error('Bad CD');
    const comp = view.getUint16(pos + 10, true);
    const nameLen = view.getUint16(pos + 28, true);
    const extraLen = view.getUint16(pos + 30, true);
    const commentLen = view.getUint16(pos + 32, true);
    const offset = view.getUint32(pos + 42, true);
    const name = new TextDecoder().decode(buf.subarray(pos + 46, pos + 46 + nameLen));
    pos += 46 + nameLen + extraLen + commentLen;

    const local = new DataView(buf.buffer, buf.byteOffset + offset, 30);
    const localNameLen = local.getUint16(26, true);
    const localExtraLen = local.getUint16(28, true);
    const dataOffset = offset + 30 + localNameLen + localExtraLen;
    const compressedSize = view.getUint32(offset + 18, true);
    const data = buf.subarray(dataOffset, dataOffset + compressedSize);
    let decompressed = data;
    if (comp === 8) decompressed = inflate(data, view.getUint32(offset + 20, true));
    else if (comp !== 0) throw new Error('Unsupported compression');
    map.set(name, decompressed);
  }
  return map;
}

function inflate(compressed, usize) {
  const out = new Uint8Array(usize);
  let op = 0, ip = 0;
  const bitbuf = { b: 0, n: 0 };
  const getbits = n => {
    while (bitbuf.n < n) {
      if (ip >= compressed.length) throw new Error('Truncated input');
      bitbuf.b |= compressed[ip++] << bitbuf.n;
      bitbuf.n += 8;
    }
    const v = bitbuf.b & ((1 << n) - 1);
    bitbuf.b >>>= n;
    bitbuf.n -= n;
    return v;
  };

  while (op < usize) {
    const bfinal = getbits(1);
    const btype = getbits(2);
    if (btype === 0) {
      bitbuf.b = bitbuf.n = 0;
      const len = compressed[ip] | (compressed[ip + 1] << 8); ip += 2;
      ip += 2;
      out.set(compressed.subarray(ip, ip + len), op);
      ip += len; op += len;
    } else if (btype === 2) {
      const hlit = getbits(5) + 257, hdist = getbits(5) + 1, hclen = getbits(4) + 4;
      const clen = new Uint8Array(19);
      const order = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
      for (let i = 0; i < hclen; i++) clen[order[i]] = getbits(3);
      const codeLens = buildHuffman(clen);
      const litLen = decodeHuffman(codeLens, hlit, getbits);
      const dist = decodeHuffman(codeLens, hdist, getbits);

      while (true) {
        const sym = decodeSym(litLen, getbits);
        if (sym < 256) {
          out[op++] = sym;
        } else if (sym === 256) {
          break;
        } else {
          let len = sym - 257;
          const lenBits = len < 8 ? 0 : Math.floor((len - 4) / 4);
          const lenAdd = lenBits ? getbits(lenBits) : 0;
          len = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258][len] + lenAdd;

          const distSym = decodeSym(dist, getbits);
          const distBits = distSym < 4 ? 0 : Math.floor((distSym - 2) / 2);
          const distAdd = distBits ? getbits(distBits) : 0;
          const distance = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577][distSym] + distAdd;

          for (let k = 0; k < len; k++) out[op + k] = out[op - distance + k];
          op += len;
        }
      }
    } else {
      throw new Error('Unsupported block type');
    }
  }
  return out;
}

function buildHuffman(lengths) {
  const max = Math.max(...lengths.filter(l => l > 0));
  const bl_count = new Uint16Array(max + 1);
  lengths.forEach(l => l > 0 && bl_count[l]++);
  const next_code = new Uint16Array(max + 1);
  let code = 0;
  for (let bits = 1; bits <= max; bits++) {
    code = (code + bl_count[bits - 1]) << 1;
    next_code[bits] = code;
  }
  const table = [];
  lengths.forEach((len, sym) => {
    if (len) table.push({ sym, len, code: next_code[len]++ });
  });
  table.sort((a, b) => a.code - b.code);
  return table;
}

function decodeHuffman(table, count, getbits) {
  const arr = new Uint8Array(count);
  for (let i = 0; i < count; ) {
    const sym = decodeSym(table, getbits);
    if (sym < 16) arr[i++] = sym;
    else if (sym === 16) {
      const rep = 3 + getbits(2);
      for (let j = 0; j < rep; j++) arr[i++] = arr[i - rep];
    } else if (sym === 17) {
      const rep = 3 + getbits(3);
      for (let j = 0; j < rep; j++) arr[i++] = 0;
    } else {
      const rep = 11 + getbits(7);
      for (let j = 0; j < rep; j++) arr[i++] = 0;
    }
  }
  return buildHuffman(arr);
}

function decodeSym(table, getbits) {
  let code = 0, len = 0;
  while (true) {
    code = (code << 1) | getbits(1); len++;
    for (const e of table) {
      if (e.len === len && e.code === code) return e.sym;
    }
  }
}

function parseSheetXml(xml) {
  const rows = [];
  const rowMatches = xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g);
  let header = null;
  for (const match of rowMatches) {
    const rowXml = match[1];
    const cellMatches = rowXml.matchAll(/<c[^>]*>(?:<v>(.*?)<\/v>|<is><t>(.*?)<\/t><\/is>)<\/c>/g);
    const cells = [];
    for (const cell of cellMatches) cells.push(cell[1] || cell[2] || '');
    if (!header) { header = cells; continue; }
    const obj = {};
    cells.forEach((v, i) => { const h = header[i] || `col${i}`; obj[h] = v; });
    rows.push(obj);
  }
  return { header, data: rows };
}

// ====================== DATA STORE ======================
let allRows = [];
let allHeaders = [];

function loadAllExcelFiles() {
  allRows = []; allHeaders = [];
  const files = fs.readdirSync(UPLOAD_DIR).filter(f => /\.(xlsx|xls)$/i.test(f));
  const headerSet = new Set();
  for (const f of files) {
    const filePath = path.join(UPLOAD_DIR, f);
    let buf;
    try { buf = fs.readFileSync(filePath); } catch { continue; }
    let parsed;
    try { parsed = parseXLSX(buf); } catch (e) { console.error(`Parse error in ${f}:`, e.message); continue; }
    parsed.header.forEach(h => headerSet.add(h));
    allRows = allRows.concat(parsed.data);
  }
  allHeaders = Array.from(headerSet);
  console.log(`Loaded ${allRows.length} rows from ${files.length} files`);
}
loadAllExcelFiles();

// ====================== HTTP SERVER ======================
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const send = (code, body, type = 'text/html') => {
    res.writeHead(code, { 'Content-Type': type });
    res.end(body);
  };
  const sendJson = obj => send(200, JSON.stringify(obj), 'application/json');

  // Static CSS
  if (req.method === 'GET' && url.pathname === '/style.css') {
    const css = fs.readFileSync(path.resolve(__dirname, 'style.css'), 'utf8');
    return send(200, css, 'text/css');
  }

  // Home Page
  if (req.method === 'GET' && url.pathname === '/') {
    const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Excel Grid</title>
<link rel="stylesheet" href="/style.css"></head><body>
<h1>Excel Grid Viewer</h1>

<div class="upload-box">
  <h3>Upload Files</h3>
  <form id="uploadForm" enctype="multipart/form-data">
    <input type="file" name="files" multiple accept=".xlsx,.xls" required>
    <button type="submit">Upload</button>
    <span id="upmsg"></span>
  </form>
</div>

<div class="upload-box">
  <h3>Load from Directory</h3>
  <form id="dirForm">
    <input type="text" id="dirPath" placeholder="e.g., C:\\data or /home/user/excel" style="width:70%;padding:8px;">
    <button type="submit">Load All</button>
    <span id="dirmsg"></span>
  </form>
</div>

<div class="controls">
  <input id="search" placeholder="Search all columns…">
  <button id="exportBtn">Export CSV</button>
  <button id="refreshBtn">Refresh</button>
</div>

<div id="summary"></div>
<div class="table-wrap"><table id="grid"><thead id="head"><tr><td>Loading...</td></tr></thead><tbody id="body"></tbody></table></div>
<div id="paging" class="paging"></div>

<script>
const qs = s => document.querySelector(s);
let page = 1, search = '';

qs('#uploadForm').onsubmit = async e => {
  e.preventDefault();
  const fd = new FormData(e.target);
  const r = await fetch('/upload', {method: 'POST', body: fd});
  const j = await r.json();
  qs('#upmsg').textContent = j.ok ? 'Uploaded ' + j.cnt : j.err;
  load();
};

qs('#dirForm').onsubmit = async e => {
  e.preventDefault();
  const dir = qs('#dirPath').value.trim();
  if (!dir) return;
  qs('#dirmsg').textContent = 'Loading...';
  const r = await fetch('/load-dir?path=' + encodeURIComponent(dir));
  const j = await r.json();
  qs('#dirmsg').textContent = j.ok ? 'Loaded ' + j.cnt + ' files' : j.err;
  load();
};

qs('#search').oninput = debounce(() => { search = qs('#search').value; page = 1; load(); }, 300);
qs('#exportBtn').onclick = () => location.href = '/export?search=' + encodeURIComponent(search);
qs('#refreshBtn').onclick = () => load();

function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }

function load() {
  fetch('/data?page=' + page + '&search=' + encodeURIComponent(search))
    .then(r => r.json())
    .then(d => {
      renderHead(d.headers);
      renderBody(d.rows);
      renderPaging(d);
      renderSummary(d.summary);
    });
}

function renderHead(h) { qs('#head').innerHTML = '<tr>' + h.map(c => '<th>' + esc(c) + '</th>').join('') + '</tr>'; }
function renderBody(rows) {
  qs('#body').innerHTML = rows.map(r => '<tr>' + allHeaders.map(c => '<td>' + esc(r[c] || '') + '</td>').join('') + '</tr>').join('') ||
    '<tr><td colspan="99">No data</td></tr>';
}
function renderPaging(d) {
  const p = d.page, tot = d.total, pages = d.pages;
  let html = '<button ' + (p===1?'disabled':'') + ' onclick="go('+(p-1)+')">Prev</button> ';
  for (let i = 1; i <= pages; i++) {
    if (i === p || i <= 2 || i > pages - 2 || Math.abs(i - p) <= 1)
      html += '<button class="' + (i===p?'active':'') + '" onclick="go(' + i + ')">' + i + '</button> ';
    else if (i === 3 || i === pages - 2) html += '… ';
  }
  html += '<button ' + (p===pages?'disabled':'') + ' onclick="go('+(p+1)+')">Next</button> ';
  html += ' (' + tot + ' rows)';
  qs('#paging').innerHTML = html;
}
function go(p) { page = p; load(); }
function renderSummary(s) {
  const div = qs('#summary');
  div.innerHTML = '<div class="sumgrid>' + Object.entries(s).map(([c, v]) =>
    '<div class="card"><b>' + esc(c) + '</b><br>Count: ' + v.cnt + ' Unique: ' + v.uni +
    (v.sum ? ' Sum: ' + v.sum.toFixed(2) : '') +
    (v.avg ? ' Avg: ' + v.avg.toFixed(2) : '') +
    (v.min != null ? ' Min: ' + v.min + ' Max: ' + v.max : '') + '</div>'
  ).join('') + '</div>';
}
function esc(t) { const d = document.createElement('div'); d.textContent = t; return d.innerHTML; }

let allHeaders = [];
fetch('/data').then(r => r.json()).then(d => allHeaders = d.headers);
load();
</script>
</body></html>`;
    return send(200, html);
  }

  // Load from Directory
  if (req.method === 'GET' && url.pathname === '/load-dir') {
    const dirPath = url.searchParams.get('path') || '';
    let realPath;
    try { realPath = path.resolve(dirPath); } catch { return sendJson({ ok: false, err: 'Invalid path' }); }
    if (!fs.existsSync(realPath) || !fs.statSync(realPath).isDirectory())
      return sendJson({ ok: false, err: 'Directory not found' });

    const files = fs.readdirSync(realPath).filter(f => /\.(xlsx|xls)$/i.test(f));
    let copied = 0;
    for (const f of files) {
      try {
        fs.copyFileSync(path.join(realPath, f), path.join(UPLOAD_DIR, f));
        copied++;
      } catch (e) { console.error(`Copy failed: ${f}`, e.message); }
    }
    loadAllExcelFiles();
    return sendJson({ ok: true, cnt: copied });
  }

  // Data API
  if (req.method === 'GET' && url.pathname === '/data') {
    let { page = 1, search = '' } = Object.fromEntries(url.searchParams);
    page = Math.max(1, +page);
    const per = 20;
    let filtered = allRows;
    if (search) {
      const q = search.toLowerCase();
      filtered = allRows.filter(r => allHeaders.some(h => (r[h] || '').toString().toLowerCase().includes(q)));
    }
    const total = filtered.length;
    const pages = Math.ceil(total / per);
    const rows = filtered.slice((page - 1) * per, page * per);

    const summary = {};
    allHeaders.forEach(h => {
      const vals = filtered.map(r => r[h]).filter(v => v != null && v !== '');
      const nums = vals.map(v => +v).filter(n => !isNaN(n));
      summary[h] = {
        cnt: vals.length,
        uni: new Set(vals.map(String)).size,
        sum: nums.reduce((a, b) => a + b, 0),
        avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
        min: nums.length ? Math.min(...nums) : null,
        max: nums.length ? Math.max(...nums) : null
      };
    });

    return sendJson({ rows, total, pages, page, headers: allHeaders, summary });
  }

  // Export CSV
  if (req.method === 'GET' && url.pathname === '/export') {
    const search = url.searchParams.get('search') || '';
    let filtered = allRows;
    if (search) {
      const q = search.toLowerCase();
      filtered = allRows.filter(r => allHeaders.some(h => (r[h] || '').toString().toLowerCase().includes(q)));
    }
    if (!filtered.length) return send(400, 'No data');
    const csv = [allHeaders.join(',')].concat(
      filtered.map(r => allHeaders.map(h => JSON.stringify(r[h] || '')).join(','))
    ).join('\r\n');
    res.writeHead(200, {
      'Content-Type': 'text/csv',
      'Content-Disposition': `attachment; filename="export-${Date.now()}.csv"`
    });
    return res.end(csv);
  }

  // Upload
  if (req.method === 'POST' && url.pathname === '/upload') {
    const boundary = req.headers['content-type'].split('boundary=')[1];
    const parts = await parseMultipart(req, boundary);
    let cnt = 0;
    for (const part of parts) {
      if (part.filename && /\.(xlsx?|xls)$/i.test(part.filename)) {
        fs.writeFileSync(path.join(UPLOAD_DIR, part.filename), part.data);
        cnt++;
      }
    }
    loadAllExcelFiles();
    return sendJson({ ok: true, cnt });
  }

  send(404, 'Not found');
});

server.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));

// ====================== MULTIPART PARSER ======================
async function parseMultipart(req, boundary) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const body = Buffer.concat(chunks);
  const bound = Buffer.from('\r\n--' + boundary);
  const parts = [];
  let start = body.indexOf(bound) + bound.length + 2;
  while (true) {
    const end = body.indexOf(bound, start);
    if (end === -1) break;
    const raw = body.subarray(start, end);
    const hdrEnd = raw.indexOf(Buffer.from('\r\n\r\n'));
    const headers = raw.subarray(0, hdrEnd).toString();
    const data = raw.subarray(hdrEnd + 4);
    const match = headers.match(/filename="([^"]*)"/);
    parts.push({ filename: match ? match[1] : null, data });
    start = end + bound.length + 2;
  }
  return parts;
}
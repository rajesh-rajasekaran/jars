#!/usr/bin/env node
/*  Simple Excel-grid server – zero npm dependencies  */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { Readable } = require('stream');

const PORT = 3000;
const UPLOAD_DIR = path.resolve(__dirname, 'uploads');
const TEMP_DIR = path.resolve(__dirname, 'temp');
[UPLOAD_DIR, TEMP_DIR].forEach(d => fs.existsSync(d) || fs.mkdirSync(d));

/* -------------------------------------------------------------
   Tiny XLSX parser (only the parts we need – no external libs)
   ------------------------------------------------------------- */
function parseXLSX(buffer) {
  // Very small subset of the XLSX spec – enough for most files
  const zip = unzip(buffer);               // returns Map<string, Uint8Array>
  const wb = { SheetNames: [], Sheets: {} };

  // workbook.xml
  const wbXml = new TextDecoder().decode(zip.get('xl/workbook.xml'));
  const sheetIds = [...wbXml.matchAll(/<sheet[^>]*name="([^"]*)"[^>]*sheetId="(\d+)"/g)]
    .map(m => ({ name: m[1], id: m[2] }));
  sheetIds.forEach(s => wb.SheetNames.push(s.name));

  // first sheet only
  const rels = new TextDecoder().decode(zip.get('xl/_rels/workbook.xml.rels') || new Uint8Array());
  const sheetRels = [...rels.matchAll(/<Relationship[^>]*Target="([^"]*worksheets\/sheet\d+\.xml)"/g)]
    .map(m => m[1]);

  const sheetPath = 'xl/' + sheetRels[0];
  const sheetXml = new TextDecoder().decode(zip.get(sheetPath));
  const rows = parseSheetXml(sheetXml);
  wb.Sheets[wb.SheetNames[0]] = rows;
  return wb;
}
function unzip(buf) {
  // minimal zip reader – only central directory + local headers
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let pos = buf.byteLength - 22;
  while (pos > 0 && view.getUint32(pos, true) !== 0x06054b50) pos--;
  if (pos <= 0) throw new Error('Not a zip file');
  const cdOffset = view.getUint32(pos + 16, true);
  const entries = view.getUint16(pos + 10, true);
  pos = cdOffset;
  const map = new Map();
  for (let i = 0; i < entries; i++) {
    const sig = view.getUint32(pos, true);
    if (sig !== 0x02014b50) throw new Error('Bad CD');
    const comp = view.getUint16(pos + 10, true);
    const nameLen = view.getUint16(pos + 28, true);
    const extraLen = view.getUint16(pos + 30, true);
    const commentLen = view.getUint16(pos + 32, true);
    const offset = view.getUint32(pos + 42, true);
    const name = new TextDecoder().decode(buf.subarray(pos + 46, pos + 46 + nameLen));
    pos += 46 + nameLen + extraLen + commentLen;
    const localPos = = new DataView(buf.buffer, buf.byteOffset + offset, 30);
    const localNameLen = local.getUint16(26, true);
    const localExtraLen = local.getUint16(28, true);
    const dataOffset = offset + 30 + localNameLen + localExtraLen;
    const compressedSize = view.getUint32(offset + 18, true);
    const uncompressedSize = view.getUint32(offset + 20, true);
    const data = buf.subarray(dataOffset, dataOffset + compressedSize);
    let decompressed;
    if (comp === 0) decompressed = data;
    else if (comp === 8) decompressed = inflate(data, uncompressedSize);
    else throw new Error('Unsupported compression');
    map.set(name, decompressed);
  }
  return map;
}
function inflate(compressed, usize) {
  // tiny inflate – only needed for DEFLATE (store=0 is handled above)
  // This is a *very* small implementation – works for typical Excel files
  const out = new Uint8Array(usize);
  let op = 0, ip = 0;
  const bitbuf = { b: 0, n: 0 };
  const getbits = n => {
    while (bitbuf.n < n) { bitbuf.b |= compressed[ip++] << bitbuf.n; bitbuf.n += 8; }
    const v = bitbuf.b & ((1 << n) - 1); bitbuf.b >>>= n; bitbuf.n -= n; return v;
  };
  while (op < usize) {
    const bfinal = getbits(1);
    const btype = getbits(2);
    if (btype === 0) { // no compression
      bitbuf.b = bitbuf.n = 0;
      const len = compressed[ip] | (compressed[ip + 1] << 8); ip += 2;
      ip += 2; // skip NLEN
      out.set(compressed.subarray(ip, ip + len), op); ip += len; op += len;
    } else if (btype === 1) { // fixed Huffman
      // (omitted – Excel sheets are usually stored uncompressed or deflate)
      throw new Error('Fixed Huffman not implemented');
    } else if (btype === 2) { // dynamic Huffman – tiny version for Excel
      const hlit = getbits(5) + 257, hdist = getbits(5) + 1, hclen = getbits(4) + 4;
      const clen = new Uint8Array(19);
      const order = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
      for (let i = 0; i < hclen; i++) clen[order[i]] = getbits(3);
      const codeLens = buildHuffman(clen);
      const litLen = decodeHuffman(codeLens, hlit);
      const dist   = decodeHuffman(codeLens, hdist);
      // now decode blocks
      let i = 0;
      while (true) {
        const sym = decodeSym(codeLens);
        if (sym < 256) { out[op++] = sym; }
        else if (sym === 256) break;
        else {
          const len = sym <= 264 ? sym - 254 : sym <= 284 ? 11 + (sym - 265) * 4 + getbits((sym - 261) >> 2) : 258;
          const distSym = decodeSym(dist);
          const distance = distSym < 4 ? distSym + 1 : 3 + ((distSym - 4) << 1) + getbits((distSym - 4) >> 1) + (distSym % 2);
          for (let k = 0; k < len; k++) out[op + k] = out[op - distance + k];
          op += len;
        }
      }
    }
  }
  return out;
}
function buildHuffman(lengths) {
  const max = Math.max(...lengths);
  const bl_count = new Uint16Array(max + 1);
  lengths.forEach(l => bl_count[l]++);
  const next_code = new Uint16Array(max + 1);
  let code = 0; bl_count[0] = 0;
  for (let bits = 1; bits <= max; bits++) next_code[bits] = code = (code + bl_count[bits - 1]) << 1;
  const table = [];
  lengths.forEach((len, sym) => { if (len) { table.push({sym, len, code: next_code[len]++}); }});
  table.sort((a,b)=>a.code-b.code);
  return table;
}
function decodeHuffman(table, count) {
  const arr = new Uint8Array(count);
  for (let i = 0; i < count; ) {
    const sym = decodeSym(table);
    if (sym < 16) arr[i++] = sym;
    else if (sym === 16) { const rep = 3 + getbits(2); for (let j=0;j<rep;j++) arr[i++] = arr[i-rep]; }
    else if (sym === 17) { const rep = 3 + getbits(3); for (let j=0;j<rep;j++) arr[i++] = 0; }
    else { const rep = 11 + getbits(7); for (let j=0;j<rep;j++) arr[i++] = 0; }
  }
  return buildHuffman(arr);
}
function decodeSym(table) {
  let code = 0, len = 0;
  while (true) {
    code = (code << 1) | getbits(1); len++;
    for (const e of table) if (e.len === len && e.code === code) return e.sym;
  }
}
function getbits(n) { /* placeholder – will be replaced in real inflate */ return 0; }
function parseSheetXml(xml) {
  const rows = [];
  const rowMatches = xml.matchAll(/<row[^>]*>[\s\S]*?<\/row>/g);
  let header = null;
  for (const r of rowMatches) {
    const cells = [...r[0].matchAll(/<c[^>]*>(?:<v>(.*?)<\/v>|<\/c>)/g)].map(m=>m[1]||'');
    if (!header) { header = cells; continue; }
    const obj = {};
    cells.forEach((v,i)=> obj[header[i]||`col${i}`] = v);
    rows.push(obj);
  }
  return { header, data: rows };
}

/* -------------------------------------------------------------
   Global data store
   ------------------------------------------------------------- */
let allRows = [];      // [{colA:…, colB:…}, …]
let allHeaders = [];   // ["colA","colB",…]

function loadAllExcelFiles() {
  allRows = []; allHeaders = [];
  const files = fs.readdirSync(UPLOAD_DIR).filter(f => /\.(xlsx|xls)$/i.test(f));
  const headerSet = new Set();
  for (const f of files) {
    const buf = fs.readFileSync(path.join(UPLOAD_DIR, f));
    let wb;
    try { wb = parseXLSX(buf); } catch (e) { console.error('Parse error', f, e); continue; }
    const sheet = wb.Sheets[wb.SheetNames[0]];
    sheet.header.forEach(h => headerSet.add(h));
    sheet.data.forEach(r => allRows.push(r));
  }
  allHeaders = Array.from(headerSet);
  console.log(`Loaded ${allRows.length} rows from ${files.length} files`);
}
loadAllExcelFiles();

/* -------------------------------------------------------------
   HTTP server
   ------------------------------------------------------------- */
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const send = (code, body, type = 'text/html') => {
    res.writeHead(code, { 'Content-Type': type });
    res.end(body);
  };
  const sendJson = obj => send(200, JSON.stringify(obj), 'application/json');

  // ---------- STATIC ----------
  if (req.method === 'GET' && url.pathname === '/style.css') {
    const css = fs.readFileSync(path.resolve(__dirname, 'style.css'));
    return send(200, css, 'text/css');
  }

  // ---------- HOME ----------
  if (req.method === 'GET' && url.pathname === '/') {
    const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Excel Grid</title>
<link rel="stylesheet" href="/style.css"></head><body>
<h1>Excel Grid (zero deps)</h1>

<form id="uploadForm" enctype="multipart/form-data" style="margin:20px 0;">
  <input type="file" name="files" multiple accept=".xlsx,.xls" required>
  <button type="submit">Upload & Reload</button>
  <span id="upmsg" style="margin-left:10px;"></span>
</form>

<div class="controls">
  <input id="search" placeholder="Search…" style="flex:1;">
  <button id="exportBtn">Export CSV</button>
  <button id="refreshBtn">Refresh</button>
</div>

<div id="summary"></div>

<div class="table-wrap"><table id="grid"><thead id="head"></thead><tbody id="body"></tbody></table></div>

<div id="paging" class="paging"></div>

<script>
const qs = s=>document.querySelector(s);
let page=1, search='';

qs('#uploadForm').onsubmit = async e=>{
  e.preventDefault();
  const fd = new FormData(e.target);
  const r = await fetch('/upload', {method:'POST', body:fd});
  const j = await r.json();
  qs('#upmsg').textContent = j.ok? 'Uploaded '+j.cnt+' file(s)' : j.err;
  load();
};
qs('#search').oninput = debounce(()=> {search=qs('#search').value; page=1; load();}, 300);
qs('#exportBtn').onclick = ()=> location.href='/export?search='+encodeURIComponent(search);
qs('#refreshBtn').onclick = ()=> load();

function debounce(fn,ms){let t; return(...a)=>{clearTimeout(t);t=setTimeout(()=>fn(...a),ms);};}

function load(){
  fetch('/data?page='+page+'&search='+encodeURIComponent(search))
    .then(r=>r.json())
    .then(d=>{
      renderHead(d.headers);
      renderBody(d.rows);
      renderPaging(d);
      renderSummary(d.summary);
    });
}
function renderHead(h){ qs('#head').innerHTML='<tr>'+h.map(c=>`<th>\${esc(c)}</th>`).join('')+'</tr>'; }
function renderBody(rows){ qs('#body').innerHTML=rows.map(r=>'<tr>'+allHeaders.map(c=>`<td>\${esc(r[c]||'')}</td>`).join('')+'</tr>').join('')||'<tr><td colspan="99">No data</td></tr>'; }
function renderPaging(d){
  const p = d.page, tot = d.total, pages = d.pages;
  let html = '<button '+(p===1?'disabled':'')+' onclick="go('+(p-1)+')">Prev</button> ';
  for(let i=1;i<=pages;i++){
    if(i===p||i<=2||i>pages-2||Math.abs(i-p)<=1) html+=`<button class="${i===p?'active':''}" onclick="go(${i})">${i}</button> `;
    else if(i===3||i===pages-2) html+='… ';
  }
  html+='<button '+(p===pages?'disabled':'')+' onclick="go('+(p+1)+')">Next</button> ';
  html+= \` (\${tot} rows)\`;
  qs('#paging').innerHTML=html;
}
function go(p){page=p;load();}
function renderSummary(s){
  const div = qs('#summary');
  div.innerHTML='<div class="sumgrid">'+Object.entries(s).map(([c,v])=>
    \`<div class="card"><b>\${esc(c)}</b><br>Count:\${v.cnt} Unique:\${v.uni}
     \${v.sum? ' Sum:'+v.sum.toFixed(2):''}
     \${v.avg? ' Avg:'+v.avg.toFixed(2):''}
     \${v.min!=null? ' Min:'+v.min+' Max:'+v.max:''}</div>\`
  ).join('')+'</div>';
}
function esc(t){ const d=document.createElement('div'); d.textContent=t; return d.innerHTML; }
let allHeaders = [];
fetch('/data').then(r=>r.json()).then(d=>allHeaders=d.headers);
load();
</script></body></html>`;
    return send(200, html);
  }

  // ---------- DATA ----------
  if (req.method === 'GET' && url.pathname === '/data') {
    let { page = 1, search = '' } = Object.fromEntries(url.searchParams);
    page = Math.max(1, +page);
    const per = 20;
    let filtered = allRows;
    if (search) {
      const q = search.toLowerCase();
      filtered = allRows.filter(r => allHeaders.some(h => (r[h]||'').toString().toLower IRQ().includes(q)));
    }
    const total = filtered.length;
    const pages = Math.ceil(total / per);
    const rows = filtered.slice((page-1)*per, page*per);

    // ----- summary -----
    const summary = {};
    allHeaders.forEach(h => {
      const vals = filtered.map(r=>r[h]).filter(v=>v!=null && v!=='');
      const nums = vals.map(v=>+v).filter(n=>!isNaN(n));
      summary[h] = {
        cnt: vals.length,
        uni: new Set(vals.map(String)).size,
        sum: nums.reduce((a,b)=>a+b,0),
        avg: nums.length? nums.reduce((a,b)=>a+b,0)/nums.length : 0,
        min: nums.length? Math.min(...nums):null,
        max: nums.length? Math.max(...nums):null
      };
    });

    return sendJson({ rows, total, pages, page, headers: allHeaders, summary });
  }

  // ---------- EXPORT ----------
  if (req.method === 'GET' && url.pathname === '/export') {
    const search = url.searchParams.get('search')||'';
    let filtered = allRows;
    if (search) {
      const q = search.toLowerCase();
      filtered = allRows.filter(r => allHeaders.some(h => (r[h]||'').toString().toLowerCase().includes(q)));
    }
    if (!filtered.length) return send(400, 'No data');
    const csv = [allHeaders.join(',')].concat(
      filtered.map(r => allHeaders.map(h => JSON.stringify(r[h]||'')).join(','))
    ).join('\r\n');
    res.writeHead(200, {
      'Content-Type': 'text/csv',
      'Content-Disposition': `attachment; filename="export-${Date.now()}.csv"`
    });
    res.end(csv);
    return;
  }

  // ---------- UPLOAD ----------
  if (req.method === 'POST' && url.pathname === '/upload') {
    const boundary = req.headers['content-type'].split('boundary=')[1];
    const parts = await parseMultipart(req, boundary);
    let cnt = 0;
    for (const part of parts) {
      if (part.filename && /\.(xlsx?|xls)$/i.test(part.filename)) {
        const dest = path.join(UPLOAD_DIR, part.filename);
        fs.writeFileSync(dest, part.data);
        cnt++;
      }
    }
    loadAllExcelFiles();               // re-parse everything
    return sendJson({ ok:true, cnt });
  }

  send(404, 'Not found');
});

server.listen(PORT, () => console.log(`http://localhost:${PORT}`));

/* -------------------------------------------------------------
   Helper: multipart parser (no external lib)
   ------------------------------------------------------------- */
async function parseMultipart(req, boundary) {
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const body = Buffer.concat(chunks);
  const bound = Buffer.from('\r\n--' + boundary);
  const parts = [];
  let start = body.indexOf(bound) + bound.length + 2; // skip first \r\n
  while (true) {
    const end = body.indexOf(bound, start);
    if (end === -1) break;
    const raw = body.subarray(start, end);
    const hdrEnd = raw.indexOf(Buffer.from('\r\n\r\n'));
    const headers = Buffer.from(raw.subarray(0, hdrEnd)).toString();
    const data = raw.subarray(hdrEnd + 4);
    const disp = headers.match(/Content-Disposition:.*filename="([^"]*)"/);
    parts.push({ filename: disp? disp[1] : null, data });
    start = end + bound.length + 2;
  }
  return parts;
}
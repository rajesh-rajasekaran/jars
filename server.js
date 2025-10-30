#!/usr/bin/env node
// server.js - Standalone HTTP Excel Merger (no external deps)

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

// === BUNDLED SheetJS (loaded once) ===
const XLSX = (function() {
  const script = fs.readFileSync(path.join(__dirname, 'xlsx.core.min.js'), 'utf8');
  return eval(script);
})();

// === CONFIG ===
const PORT = 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');

// === HELPERS ===
function sendFile(res, filePath, contentType = 'text/html') {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not Found');
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}

function escapeHtml(text) {
  return String(text).replace(/[&<>"']/g, m => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[m]);
}

// === SERVER ===
const server = http.createServer(async (req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  // Serve static files
  if (pathname === '/' || pathname === '/index.html') {
    return sendFile(res, path.join(PUBLIC_DIR, 'index.html'));
  }

  if (pathname === '/xlsx.core.min.js') {
    return sendFile(res, path.join(__dirname, 'xlsx.core.min.js'), 'application/javascript');
  }

  // API: Process folder
  if (pathname === '/api/process' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const { files } = JSON.parse(body);
        if (!Array.isArray(files)) throw new Error('Invalid files');

        const merged = [];
        const fileColors = {};
        const colors = ['#e3f2fd','#f3e5f5','#e8f5e9','#fff3e0','#fce4ec','#e0f7fa','#f1f8e9','#fff8e1','#e1f5fe'];
        let processed = 0;

        res.writeHead(200, {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive'
        });

        const sendProgress = (msg, percent) => {
          res.write(`data: ${JSON.stringify({ type: 'progress', message: msg, percent })}\n\n`);
        };

        for (const fileObj of files) {
          processed++;
          sendProgress(`Reading ${fileObj.name}...`, (processed / files.length) * 100);

          const arrayBuffer = Uint8Array.from(atob(fileObj.data.split(',')[1]), c => c.charCodeAt(0)).buffer;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

          if (json.length < 1) continue;

          const fileName = fileObj.name;
          if (!fileColors[fileName]) {
            fileColors[fileName] = colors[Object.keys(fileColors).length % colors.length];
          }

          const headers = json[0];
          json.slice(1).forEach(row => {
            const obj = { 'Source File': fileName, 'Sheet': sheetName };
            headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
            merged.push(obj);
          });
        }

        sendProgress('Finalizing...', 100);
        setTimeout(() => {
          res.write(`data: ${JSON.stringify({ type: 'done', data: merged, colors: fileColors })}\n\n`);
          res.end();
        }, 300);
      } catch (err) {
        res.write(`data: ${JSON.stringify({ type: 'error', message: err.message })}\n\n`);
        res.end();
      }
    });
    return;
  }

  // API: Export CSV
  if (pathname === '/api/export' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { data } = JSON.parse(body);
        const headers = Object.keys(data[0] || {});
        const csv = [
          headers.map(h => `"${h}"`).join(','),
          ...data.map(row => headers.map(h => `"${(row[h]+'').replace(/"/g, '""')}"`).join(','))
        ].join('\n');

        res.writeHead(200, {
          'Content-Type': 'text/csv',
          'Content-Disposition': `attachment; filename="merged-${new Date().toISOString().slice(0,19).replace(/:/g,'-')}.csv"`
        });
        res.end('\uFEFF' + csv);
      } catch (err) {
        res.writeHead(500);
        res.end('Export failed');
      }
    });
    return;
  }

  res.writeHead(404);
  res.end('Not Found');
});

// === START ===
server.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`Open your browser and go to the URL above.`);
});
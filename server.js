#!/usr/bin/env node
// server.js - HTTP Excel Merger (CDN SheetJS)

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

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

// === SERVER ===
const server = http.createServer((req, res) => {
  const parsed = url.parse(req.url, true);
  const pathname = parsed.pathname;

  // Serve index
  if (pathname === '/' || pathname === '/index.html') {
    return sendFile(res, path.join(PUBLIC_DIR, 'index.html'));
  }

  // === API: /api/process (SSE) ===
  if (pathname === '/api/process' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk.toString());
    req.on('end', () => {
      try {
        const { files } = JSON.parse(body);
        if (!Array.isArray(files)) throw new Error('Invalid files');

        res.writeHead(200, {
          'Content-Type': 'text/event-stream',
          'Cache-Control': 'no-cache',
          'Connection': 'keep-alive',
          'Access-Control-Allow-Origin': '*'
        });

        const merged = [];
        const fileColors = {};
        const colors = ['#e3f2fd','#f3e5f5','#e8f5e9','#fff3e0','#fce4ec','#e0f7fa','#f1f8e9','#fff8e1','#e1f5fe'];
        let processed = 0;

        const send = (type, data) => {
          res.write(`data: ${JSON.stringify({ type, ...data })}\n\n`);
        };

        (async () => {
          for (const fileObj of files) {
            processed++;
            send('progress', {
              message: `Reading ${fileObj.name}...`,
              percent: (processed / files.length) * 100
            });

            const base64 = fileObj.data.split(',')[1];
            const buffer = Buffer.from(base64, 'base64');
            const workbook = XLSX.read(buffer, { type: 'buffer' });
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

          send('progress', { message: 'Finalizing...', percent: 100 });
          await new Promise(r => setTimeout(r, 300));
          send('done', { data: merged, colors: fileColors });
          res.end();
        })();
      } catch (err) {
        res.write(`data: ${JSON.stringify({ type: 'error', message: err.message })}\n\n`);
        res.end();
      }
    });
    return;
  }

  // === API: /api/export ===
  if (pathname === '/api/export' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try {
        const { data } = JSON.parse(body);
        if (!Array.isArray(data) || data.length === 0) throw new Error('No data');

        const headers = Object.keys(data[0]);
        const csv = [
          headers.map(h => `"${h}"`).join(','),
          ...data.map(row => headers.map(h => `"${(row[h]+'').replace(/"/g, '""')}"`).join(','))
        ].join('\n');

        res.writeHead(200, {
          'Content-Type': 'text/csv; charset=utf-8',
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

server.listen(PORT, '127.0.0.1', () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Merged Multi-Excel Reader</title>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --color-1: #e3f2fd; --color-2: #f3e5f5; --color-3: #e8f5e9; 
      --color-4: #fff3e0; --color-5: #fce4ec; --color-6: #e0f7fa;
      --color-7: #f1f8e9; --color-8: #fff8e1; --color-9: #e1f5fe;
    }
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      background: #f0f4f8;
      color: #2c3e50;
    }
    .container {
      max-width: 1200px;
      margin: 30px auto;
      background: white;
      border-radius: 12px;
      box-shadow: 0 8px 25px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    header {
      background: #2c3e50;
      color: white;
      padding: 20px;
      text-align: center;
    }
    h1 { margin: 0; font-size: 1.8em; }
    .subtitle { margin: 8px 0 0; font-weight: normal; opacity: 0.9; }

    .controls {
      padding: 20px;
      background: #f8f9fa;
      border-bottom: 1px solid #dee2e6;
      display: flex;
      flex-wrap: wrap;
      gap: 15px;
      align-items: center;
    }
    input[type="file"] {
      flex: 1;
      min-width: 250px;
      padding: 10px;
      border: 2px dashed #95a5a6;
      border-radius: 8px;
      background: #ecf0f1;
    }
    .btn {
      padding: 10px 18px;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-weight: bold;
      transition: 0.2s;
    }
    .btn-clear {
      background: #e74c3c;
      color: white;
    }
    .btn-clear:hover { background: #c0392b; }

    .search-container {
      padding: 15px 20px;
      background: #f8f9fa;
      border-bottom: 1px solid #dee2e6;
    }
    .search-box {
      width: 100%;
      padding: 12px;
      border: 1px solid #ced4da;
      border-radius: 8px;
      font-size: 1em;
    }

    .summary {
      padding: 0 20px 15px;
      font-weight: bold;
      color: #2980b9;
      font-size: 1.1em;
    }

    .legend {
      padding: 15px 20px;
      background: #f1f3f5;
      border-bottom: 1px solid #dee2e6;
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      font-size: 0.9em;
    }
    .legend-item {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .legend-color {
      width: 16px;
      height: 16px;
      border-radius: 4px;
      border: 1px solid #ccc;
    }

    .table-container {
      padding: 20px;
      max-height: 600px;
      overflow: auto;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.95em;
    }
    th, td {
      border: 1px solid #dee2e6;
      padding: 10px;
      text-align: left;
    }
    th {
      background-color: #2c3e50;
      color: white;
      position: sticky;
      top: 0;
      z-index: 10;
    }
    tr:nth-child(even) td { background-color: #f9f9fb; }
    .no-data {
      text-align: center;
      color: #7f8c8d;
      font-style: italic;
      padding: 40px;
      font-size: 1.1em;
    }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>Merged Multi-Excel Reader</h1>
      <p class="subtitle">Upload multiple Excel files → See all data in one table with color coding</p>
    </header>

    <div class="controls">
      <input type="file" id="excelInput" accept=".xlsx,.xls" multiple />
      <button class="btn btn-clear" id="clearAll">Clear All</button>
    </div>

    <div class="search-container">
      <input type="text" id="searchInput" class="search-box" placeholder="Search across all files and sheets..." />
    </div>

    <div id="summary" class="summary"></div>
    <div id="legend" class="legend"></div>

    <div class="table-container">
      <div id="output"></div>
    </div>
  </div>

  <script>
    const fileInput = document.getElementById('excelInput');
    const output = document.getElementById('output');
    const searchInput = document.getElementById('searchInput');
    const summary = document.getElementById('summary');
    const legend = document.getElementById('legend');
    const clearAllBtn = document.getElementById('clearAll');

    let mergedData = []; // Array of { row: [...], fileName, sheetName }
    let filteredData = [];
    const fileColors = {};
    const colorPool = [
      'var(--color-1)', 'var(--color-2)', 'var(--color-3)',
      'var(--color-4)', 'var(--color-5)', 'var(--color-6)',
      'var(--color-7)', 'var(--color-8)', 'var(--color-9)'
    ];

    fileInput.addEventListener('change', function(e) {
      const files = Array.from(e.target.files);
      if (files.length === 0) return;

      let loadedCount = 0;
      files.forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetNames = workbook.SheetNames;

          // Assign color if new file
          if (!fileColors[file.name]) {
            const colorIndex = Object.keys(fileColors).length % colorPool.length;
            fileColors[file.name] = colorPool[colorIndex];
          }

          sheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            rows.forEach((row, idx) => {
              if (idx === 0) return; // Skip header
              mergedData.push({
                row: [...row],
                fileName: file.name,
                sheetName: sheetName
              });
            });
          });

          loadedCount++;
          if (loadedCount === files.length) {
            filteredData = [...mergedData];
            renderAll();
          }
        };
        reader.readAsArrayBuffer(file);
      });

      fileInput.value = '';
    });

    clearAllBtn.addEventListener('click', () => {
      if (confirm('Remove all data and clear the table?')) {
        mergedData = [];
        filteredData = [];
        fileColors = {};
        renderAll();
      }
    });

    searchInput.addEventListener('input', () => {
      const query = searchInput.value.trim().toLowerCase();
      if (query === '') {
        filteredData = [...mergedData];
      } else {
        filteredData = mergedData.filter(item => {
          return item.row.some(cell => 
            cell != null && cell.toString().toLowerCase().includes(query)
          ) || item.fileName.toLowerCase().includes(query)
            || item.sheetName.toLowerCase().includes(query);
        });
      }
      renderTable();
      updateSummary();
    });

    function renderAll() {
      renderLegend();
      renderTable();
      updateSummary();
    }

    function renderLegend() {
      legend.innerHTML = '';
      if (Object.keys(fileColors).length === 0) {
        legend.innerHTML = '<i>No files loaded</i>';
        return;
      }
      Object.entries(fileColors).forEach(([fileName, color]) => {
        const item = document.createElement('div');
        item.className = 'legend-item';
        item.innerHTML = `
          <div class="legend-color" style="background-color: ${color};"></div>
          <span>${escapeHtml(fileName)}</span>
        `;
        legend.appendChild(item);
      });
    }

    function renderTable() {
      if (filteredData.length === 0) {
        output.innerHTML = '<p class="no-data">No data loaded. Upload Excel files to begin.</p>';
        return;
      }

      // Build master header from all rows (union of columns)
      const allHeaders = new Set();
      allHeaders.add('Source File');
      allHeaders.add('Sheet');
      filteredData.forEach(item => {
        item.row.forEach((_, i) => allHeaders.add(`Col ${i + 1}`));
      });
      const headers = Array.from(allHeaders);

      let html = '<table><thead><tr>';
      headers.forEach(h => html += `<th>${escapeHtml(h)}</th>`);
      html += '</tr></thead><tbody>';

      filteredData.forEach(item => {
        const color = fileColors[item.fileName] || '#fff';
        html += `<tr style="background-color: ${color};">`;
        html += `<td><strong>${escapeHtml(item.fileName)}</strong></td>`;
        html += `<td>${escapeHtml(item.sheetName)}</td>`;
        // Fill data columns
        item.row.forEach((cell, i) => {
          html += `<td>${escapeHtml(cell)}</td>`;
        });
        // Fill empty cells for missing columns
        const dataCols = item.row.length;
        for (let i = dataCols; i < headers.length - 2; i++) {
          html += `<td></td>`;
        }
        html += '</tr>';
      });

      html += '</tbody></table>';
      output.innerHTML = html;
    }

    function updateSummary() {
      const total = mergedData.length;
      const shown = filteredData.length;
      summary.textContent = `Showing ${shown} of ${total} total records across all files`;
    }

    function escapeHtml(text) {
      if (text == null) return '';
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
  </script>
</body>
</html>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Excel File Reader with Search</title>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 40px;
      background: #f4f4f4;
    }
    h1 { color: #333; }
    .container {
      max-width: 1000px;
      margin: auto;
      background: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 0 10px rgba(0,0,0,0.1);
    }
    input[type="file"], input[type="text"] {
      margin: 10px 0;
      padding: 8px;
      width: 100%;
      box-sizing: border-box;
    }
    .search-container {
      margin: 15px 0;
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .search-container input {
      flex: 1;
    }
    .summary {
      font-weight: bold;
      color: #007bff;
      margin: 10px 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 10px;
    }
    th, td {
      border: 1px solid #ccc;
      padding: 8px;
      text-align: left;
    }
    th {
      background-color: #f2f2f2;
    }
    .sheet-tabs {
      margin: 15px 0;
    }
    .sheet-tab {
      display: inline-block;
      padding: 8px 12px;
      margin-right: 5px;
      background: #eee;
      border: 1px solid #ccc;
      border-radius: 4px;
      cursor: pointer;
    }
    .sheet-tab.active {
      background: #007bff;
      color: white;
      border-color: #007bff;
    }
    .no-data {
      color: #888;
      font-style: italic;
      text-align: center;
      padding: 20px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Excel File Reader</h1>
    <p>Upload an Excel file (.xlsx or .xls) to view and search its contents.</p>
    <input type="file" id="excelInput" accept=".xlsx,.xls" />

    <div id="sheetTabs" class="sheet-tabs"></div>

    <div class="search-container">
      <input type="text" id="searchInput" placeholder="Search in current sheet..." />
    </div>
    <div id="summary" class="summary"></div>

    <div id="output"></div>
  </div>

  <script>
    const fileInput = document.getElementById('excelInput');
    const output = document.getElementById('output');
    const sheetTabs = document.getElementById('sheetTabs');
    const searchInput = document.getElementById('searchInput');
    const summary = document.getElementById('summary');

    let currentWorkbook = null;
    let currentSheetName = '';
    let allRows = []; // Stores header + all data rows
    let filteredRows = [];

    fileInput.addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });

        sheetTabs.innerHTML = '';
        output.innerHTML = '';
        searchInput.value = '';
        summary.textContent = '';

        const sheetNames = currentWorkbook.SheetNames;

        sheetNames.forEach((sheetName, index) => {
          const tab = document.createElement('div');
          tab.className = 'sheet-tab';
          if (index === 0) tab.classList.add('active');
          tab.textContent = sheetName;
          tab.onclick = () => switchSheet(sheetName, tab);
          sheetTabs.appendChild(tab);
        });

        if (sheetNames.length > 0) {
          switchSheet(sheetNames[0], sheetTabs.children[0]);
        }
      };
      reader.readAsArrayBuffer(file);
    });

    function switchSheet(sheetName, activeTab) {
      currentSheetName = sheetName;

      // Update active tab
      document.querySelectorAll('.sheet-tab').forEach(tab => tab.classList.remove('active'));
      activeTab.classList.add('active');

      const worksheet = currentWorkbook.Sheets[sheetName];
      allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (allRows.length === 0) {
        output.innerHTML = '<p class="no-data">No data in this sheet.</p>';
        summary.textContent = '';
        return;
      }

      filteredRows = [...allRows];
      renderTable();
      updateSummary();
    }

    searchInput.addEventListener('input', function() {
      const query = searchInput.value.trim().toLowerCase();
      if (!allRows.length) return;

      if (query === '') {
        filteredRows = allRows;
      } else {
        filteredRows = allRows.filter((row, index) => {
          if (index === 0) return true; // Always show header
          return row.some(cell => 
            cell != null && cell.toString().toLowerCase().includes(query)
          );
        });
      }

      renderTable();
      updateSummary();
    });

    function renderTable() {
      if (filteredRows.length === 0) {
        output.innerHTML = '<p class="no-data">No matching records found.</p>';
        return;
      }

      let html = '<table><thead><tr>';
      const header = filteredRows[0];
      header.forEach(col => {
        html += `<th>${escapeHtml(col)}</th>`;
      });
      html += '</tr></thead><tbody>';

      filteredRows.slice(1).forEach(row => {
        html += '<tr>';
        row.forEach(cell => {
          html += `<td>${escapeHtml(cell)}</td>`;
        });
        html += '</tr>';
      });

      html += '</tbody></table>';
      output.innerHTML = html;
    }

    function updateSummary() {
      const total = allRows.length > 0 ? allRows.length - 1 : 0;
      const shown = filteredRows.length > 0 ? filteredRows.length - 1 : 0;
      summary.textContent = `Showing ${shown} of ${total} records`;
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
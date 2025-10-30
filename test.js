<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Excel File Reader</title>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 40px;
      background: #f4f4f4;
    }
    h1 { color: #333; }
    .container {
      max-width: 900px;
      margin: auto;
      background: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 0 10px rgba(0,0,0,0.1);
    }
    input[type="file"] {
      margin: 15px 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
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
  </style>
</head>
<body>
  <div class="container">
    <h1>Excel File Reader</h1>
    <p>Upload an Excel file (.xlsx or .xls) to view its contents.</p>
    <input type="file" id="excelInput" accept=".xlsx,.xls" />
    
    <div id="sheetTabs" class="sheet-tabs"></div>
    <div id="output"></div>
  </div>

  <script>
    const fileInput = document.getElementById('excelInput');
    const output = document.getElementById('output');
    const sheetTabs = document.getElementById('sheetTabs');

    fileInput.addEventListener('change', function(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload =发出 function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Clear previous content
        sheetTabs.innerHTML = '';
        output.innerHTML = '';

        const sheetNames = workbook.SheetNames;

        // Create tabs for each sheet
        sheetNames.forEach((sheetName, index) => {
          const tab = document.createElement('div');
          tab.className = 'sheet-tab';
          if (index === 0) tab.classList.add('active');
          tab.textContent = sheetName;
          tab.onclick = () => displaySheet(workbook, sheetName, tab);
          sheetTabs.appendChild(tab);
        });

        // Display first sheet by default
        if (sheetNames.length > 0) {
          displaySheet(workbook, sheetNames[0], sheetTabs.children[0]);
        }
      };

      reader.readAsArrayBuffer(file);
    });

    function displaySheet(workbook, sheetName, activeTab) {
      // Update active tab
      document.querySelectorAll('.sheet-tab').forEach(tab => tab.classList.remove('active'));
      activeTab.classList.add('active');

      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Create table
      let html = '<table><thead><tr>';
      if (jsonData.length > 0) {
        jsonData[0].forEach(header => {
          html += `<th>${escapeHtml(header)}</th>`;
        });
      }
      html += '</tr></thead><tbody>';

      jsonData.slice(1).forEach(row => {
        html += '<tr>';
        row.forEach(cell => {
          html += `<td>${escapeHtml(cell)}</td>`;
        });
        html += '</tr>';
      });

      html += '</tbody></table>';
      output.innerHTML = html;
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
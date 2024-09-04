document.addEventListener('DOMContentLoaded', function () {
  loadSavedTheme();

  document.getElementById('upload').addEventListener('change', handleFiles, false);
  document.getElementById('export').addEventListener('click', exportToExcel, false);
});

const allData = [];

function handleFiles(e) {
  const files = e.target.files;
  const promises = [];

  Array.from(files).forEach(file => {
      const reader = new FileReader();
      promises.push(new Promise((resolve, reject) => {
          reader.onload = function(event) {
              const data = new Uint8Array(event.target.result);
              const workbook = XLSX.read(data, { type: 'array' });

              workbook.SheetNames.forEach(sheetName => {
                  const sheet = workbook.Sheets[sheetName];
                  const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                  allData.push(...sheetData);
              });
              resolve();
          };
          reader.onerror = reject;
          reader.readAsArrayBuffer(file);
      }));
  });

  Promise.all(promises).then(() => {
      renderTable(allData);
      document.getElementById('export').style.display = 'block';
  });
}

function renderTable(data) {
  const table = document.getElementById('output-table');
  table.innerHTML = '';  // Clear previous table content

  data.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      row.forEach(cell => {
          const cellElement = rowIndex === 0 ? document.createElement('th') : document.createElement('td');
          cellElement.textContent = cell || '';  // Handle empty cells
          tr.appendChild(cellElement);
      });
      table.appendChild(tr);
  });
}

function exportToExcel() {
  const ws = XLSX.utils.aoa_to_sheet(allData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Unificado");
  XLSX.writeFile(wb, "Planilhas_Unificadas.xlsx");
}

function toggleTheme() {
  const body = document.body;
  const themeToggle = document.querySelector('.theme-toggle');
  if (body.classList.contains('light-mode')) {
      body.classList.remove('light-mode');
      body.classList.add('dark-mode');
      themeToggle.textContent = '🌙';
      localStorage.setItem('theme', 'dark-mode');
  } else {
      body.classList.remove('dark-mode');
      body.classList.add('light-mode');
      themeToggle.textContent = '🌞';
      localStorage.setItem('theme', 'light-mode');
  }
}

function loadSavedTheme() {
  const savedTheme = localStorage.getItem('theme');
  const body = document.body;
  const themeToggle = document.querySelector('.theme-toggle');

  if (savedTheme) {
      body.classList.add(savedTheme);
      themeToggle.textContent = savedTheme === 'dark-mode' ? '🌙' : '🌞';
  } else {
      body.classList.add('light-mode');
      themeToggle.textContent = '🌞';
  }
}

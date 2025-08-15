let spreadsheetData = {};
let columnHeaders = {};
let analysisResults = {};
let currentMappings = {};
let currentNumColumns = 3;

function handleFile(fileNumber, input) {
  const file = input.files[0];
  if (!file) return;

  const status = document.getElementById(`status${fileNumber}`);
  status.innerHTML = '<span class="text-blue-600">Carregando...</span>';

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

      if (jsonData.length > 0) {
        spreadsheetData[fileNumber] = jsonData;
        columnHeaders[fileNumber] = jsonData[0];
        status.innerHTML = `<span class="text-green-600">✓ ${jsonData.length - 1} registros</span>`;

        updateColumnSection();
      }
    } catch (error) {
      status.innerHTML = '<span class="text-red-600">Erro ao ler arquivo</span>';
    }
  };
  reader.readAsArrayBuffer(file);
}

function updateColumnSection() {
  const loadedFiles = Object.keys(spreadsheetData);
  if (loadedFiles.length < 2) return;

  document.getElementById('columnSection').classList.remove('hidden');
  updateColumnSelectors();
}

function updateColumnSelectors() {
  const loadedFiles = Object.keys(spreadsheetData);
  const numColumns = parseInt(document.getElementById('numColumns').value) || 3;

  const mappingsDiv = document.getElementById('columnMappings');
  mappingsDiv.innerHTML = '';

  loadedFiles.forEach(fileNum => {
    const headers = columnHeaders[fileNum];
    const div = document.createElement('div');
    div.className = 'border border-gray-200 rounded-lg p-4';

    let columnsHtml = '';
    for (let i = 0; i < numColumns; i++) {
      columnsHtml += `
                        <div>
                            <label class="block text-sm font-medium text-gray-700 mb-1">Coluna ${i + 1} para comparar:</label>
                            <select id="column_${fileNum}_${i}" class="w-full border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
                                <option value="">Selecione a coluna...</option>
                                ${headers.map((header, index) => `<option value="${index}">${header}</option>`).join('')}
                            </select>
                        </div>
                    `;
    }

    div.innerHTML = `
                    <h3 class="font-semibold text-gray-800 mb-3">Planilha ${fileNum}</h3>
                    <div class="grid grid-cols-1 md:grid-cols-${Math.min(numColumns, 3)} gap-4">
                        ${columnsHtml}
                    </div>
                `;

    mappingsDiv.appendChild(div);
  });
}

function analyzeData() {
  const loadedFiles = Object.keys(spreadsheetData);
  const numColumns = parseInt(document.getElementById('numColumns').value) || 3;
  const results = {};

  // Get column mappings for each file
  const mappings = {};
  loadedFiles.forEach(fileNum => {
    mappings[fileNum] = {};
    for (let i = 0; i < numColumns; i++) {
      const columnSelect = document.getElementById(`column_${fileNum}_${i}`);
      if (columnSelect && columnSelect.value !== '') {
        mappings[fileNum][`column_${i}`] = columnSelect.value;
      }
    }
  });

  // Process each spreadsheet
  loadedFiles.forEach(fileNum => {
    const data = spreadsheetData[fileNum];
    const mapping = mappings[fileNum];
    results[fileNum] = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const record = {
        columns: {},
        originalRow: row,
        isUnique: false
      };

      // Get data for each selected column
      for (let j = 0; j < numColumns; j++) {
        const columnKey = `column_${j}`;
        if (mapping[columnKey] !== undefined) {
          record.columns[columnKey] = row[mapping[columnKey]] || '';
        }
      }

      results[fileNum].push(record);
    }
  });

  // Compare records between spreadsheets
  loadedFiles.forEach(fileNum => {
    results[fileNum].forEach(record => {
      let foundInOther = false;

      loadedFiles.forEach(otherFileNum => {
        if (fileNum === otherFileNum) return;

        const found = results[otherFileNum].some(otherRecord => {
          // Check if all selected columns match
          for (let i = 0; i < numColumns; i++) {
            const columnKey = `column_${i}`;
            const value1 = record.columns[columnKey];
            const value2 = otherRecord.columns[columnKey];

            if (value1 && value2) {
              if (value1.toLowerCase().trim() !== value2.toLowerCase().trim()) {
                return false;
              }
            }
          }
          return true;
        });

        if (found) foundInOther = true;
      });

      record.isUnique = !foundInOther;
    });
  });

  // Save results for download
  analysisResults = results;
  currentMappings = mappings;
  currentNumColumns = numColumns;

  displayResults(results, mappings, numColumns);
}

function displayResults(results, mappings, numColumns) {
  document.getElementById('resultsSection').classList.remove('hidden');
  const resultsDiv = document.getElementById('comparisonResults');
  resultsDiv.innerHTML = '';

  Object.keys(results).forEach(fileNum => {
    const data = results[fileNum];
    const mapping = mappings[fileNum];

    const div = document.createElement('div');
    div.className = 'border border-gray-200 rounded-lg overflow-hidden';

    const uniqueCount = data.filter(record => record.isUnique).length;

    // Create table headers dynamically
    let tableHeaders = '';
    for (let i = 0; i < numColumns; i++) {
      const columnKey = `column_${i}`;
      if (mapping[columnKey] !== undefined) {
        const headerName = columnHeaders[fileNum][mapping[columnKey]];
        tableHeaders += `<th class="px-4 py-3 text-left text-sm font-medium text-gray-700">${headerName}</th>`;
      }
    }

    // Create table rows dynamically
    const tableRows = data.map(record => {
      let rowCells = '';
      for (let i = 0; i < numColumns; i++) {
        const columnKey = `column_${i}`;
        const value = record.columns[columnKey] || '-';
        rowCells += `<td class="px-4 py-3 text-sm ${record.isUnique ? 'text-blue-900 font-medium' : 'text-gray-900'}">${value}</td>`;
      }

      return `
                        <tr class="${record.isUnique ? 'bg-blue-50 border-l-4 border-blue-400' : 'hover:bg-gray-50'}">
                            ${rowCells}
                            <td class="px-4 py-3 text-sm">
                                ${record.isUnique ?
          '<span class="bg-blue-100 text-blue-800 px-2 py-1 rounded-full text-xs font-medium">Único</span>' :
          '<span class="bg-green-100 text-green-800 px-2 py-1 rounded-full text-xs font-medium">Encontrado</span>'
        }
                            </td>
                        </tr>
                    `;
    }).join('');

    div.innerHTML = `
                    <div class="bg-gray-50 px-6 py-4 border-b">
                        <h3 class="text-lg font-semibold text-gray-800">Planilha ${fileNum}</h3>
                        <p class="text-sm text-gray-600">${uniqueCount} registros únicos encontrados</p>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full">
                            <thead class="bg-gray-100">
                                <tr>
                                    ${tableHeaders}
                                    <th class="px-4 py-3 text-left text-sm font-medium text-gray-700">Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${tableRows}
                            </tbody>
                        </table>
                    </div>
                `;

    resultsDiv.appendChild(div);
  });

  // Scroll to results
  document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });
}

function downloadUniqueOnly() {
  if (!analysisResults || Object.keys(analysisResults).length === 0) {
    alert('Nenhum resultado para baixar. Execute a análise primeiro.');
    return;
  }

  const workbook = XLSX.utils.book_new();

  Object.keys(analysisResults).forEach(fileNum => {
    const data = analysisResults[fileNum];
    const mapping = currentMappings[fileNum];
    const uniqueRecords = data.filter(record => record.isUnique);

    if (uniqueRecords.length === 0) return;

    // Create headers
    const headers = [];
    for (let i = 0; i < currentNumColumns; i++) {
      const columnKey = `column_${i}`;
      if (mapping[columnKey] !== undefined) {
        headers.push(columnHeaders[fileNum][mapping[columnKey]]);
      }
    }
    headers.push('Status');

    // Create data rows
    const rows = [headers];
    uniqueRecords.forEach(record => {
      const row = [];
      for (let i = 0; i < currentNumColumns; i++) {
        const columnKey = `column_${i}`;
        row.push(record.columns[columnKey] || '');
      }
      row.push('Único');
      rows.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, `Planilha ${fileNum} - Únicos`);
  });

  XLSX.writeFile(workbook, 'registros_unicos.xlsx');
}

function downloadAllResults() {
  if (!analysisResults || Object.keys(analysisResults).length === 0) {
    alert('Nenhum resultado para baixar. Execute a análise primeiro.');
    return;
  }

  const workbook = XLSX.utils.book_new();

  Object.keys(analysisResults).forEach(fileNum => {
    const data = analysisResults[fileNum];
    const mapping = currentMappings[fileNum];

    // Create headers
    const headers = [];
    for (let i = 0; i < currentNumColumns; i++) {
      const columnKey = `column_${i}`;
      if (mapping[columnKey] !== undefined) {
        headers.push(columnHeaders[fileNum][mapping[columnKey]]);
      }
    }
    headers.push('Status');

    // Create data rows
    const rows = [headers];
    data.forEach(record => {
      const row = [];
      for (let i = 0; i < currentNumColumns; i++) {
        const columnKey = `column_${i}`;
        row.push(record.columns[columnKey] || '');
      }
      row.push(record.isUnique ? 'Único' : 'Encontrado');
      rows.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, `Planilha ${fileNum} - Completa`);
  });

  XLSX.writeFile(workbook, 'comparacao_completa.xlsx');
}
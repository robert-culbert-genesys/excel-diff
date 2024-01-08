document.getElementById('fileInput').addEventListener('change', handleFileSelect);
document.getElementById('sheet1').addEventListener('change', compareWorksheets);
document.getElementById('sheet2').addEventListener('change', compareWorksheets);

let file;
let sheetNames;
let workbook;

function handleFileSelect(event) {
    file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });

            sheetNames = workbook.SheetNames;

            // Update the dropdowns with sheet names
            updateDropdown('sheet1', sheetNames);
            updateDropdown('sheet2', sheetNames);

            // Show the select elements when a file is selected
            document.getElementById('fileInputs').style.display = 'block';

            // Set defaults for drop down
            const dropdown1 = document.getElementById('sheet1');
            dropdown1.selectedIndex = sheetNames.length - 2;

            const dropdown2 = document.getElementById('sheet2');
            dropdown2.selectedIndex = sheetNames.length - 1;

            compareWorksheets();
        };
        reader.readAsArrayBuffer(file);
    }
}

function updateDropdown(id, sheetNames) {
    const dropdown = document.getElementById(id);
    dropdown.innerHTML = '';

    for (let i = 0; i < sheetNames.length; i++) {
        const sheetName = sheetNames[i];
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        dropdown.appendChild(option);
    }
}

function compareWorksheets() {
    const sheet1Name = document.getElementById('sheet1').value;
    const sheet2Name = document.getElementById('sheet2').value;

    const sheet1 = workbook.Sheets[sheet1Name];
    const sheet2 = workbook.Sheets[sheet2Name];

    const sheet1Data = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
    const sheet2Data = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

    compareLogic(file.name, sheet1Name, sheet2Name, sheet1Data, sheet2Data);
}

function compareLogic(excelFile, sheet1Name, sheet2Name, sheet1Data, sheet2Data) {
    // Identify rows removed from sheet1
    const removedFromSheet1 = sheet1Data.filter(row => !sheet2Data.some(otherRow => row[0] === otherRow[0]));

    // Identify rows added to sheet2
    const addedToSheet2 = sheet2Data.filter(row => !sheet1Data.some(otherRow => row[0] === otherRow[0]));

    // Identify rows with changes in data
    const changedEntries = sheet1Data
        .filter(row => sheet2Data.some(otherRow => row[0] === otherRow[0]))
        .filter(row => {
            const correspondingRow = sheet2Data.find(otherRow => row[0] === otherRow[0]);
            return !Object.entries(row).every(([key, value]) => correspondingRow[key] === value);
        });



    clearExisting();
    createTable(removedFromSheet1, `Employees only present in Sheet: ${sheet1Name}`, sheet1Data);
    createTable(addedToSheet2, `Employees only present in Sheet: ${sheet2Name}`, sheet1Data);
    printChangedEntries(changedEntries, sheet1Data, sheet2Data, sheet1Data[0]);
}

function clearExisting() {
    const existingTables = document.querySelectorAll('#comparisonTable');
    const existingHeadings = document.querySelectorAll('h2');

    for (let i = 0; i < existingTables.length; i++) {
        existingTables[i].parentNode.removeChild(existingTables[i]);
    }

    for (let i = 0; i < existingHeadings.length; i++) {
        existingHeadings[i].parentNode.removeChild(existingHeadings[i]);
    }
}

function printChangedEntries(changedEntries, sheet1Data, sheet2Data, headerRow) {
    const heading = document.createElement('h2');
    heading.textContent = 'Modified Employees:';
    document.body.appendChild(heading);

    for (let i = 0; i < changedEntries.length; i++) {
        const row = changedEntries[i];
        const correspondingRow = sheet2Data.find(otherRow => row[0] === otherRow[0]);

        // Find and print the differences between the two rows
        const differences = Object.entries(row).filter(([key, value]) => correspondingRow[key] !== value);

        const difference = `${row.slice(0, 3).join(' ')}: ${differences.map(([key, value]) => `Column: ${headerRow[key]} changed from ${sheet1Data.find(sheet1Row => sheet1Row[0] === row[0])[key]} to ${correspondingRow[key]}`).join(', ')}`;

        const heading = document.createElement('p');
        heading.id = 'comparisonTable'; // Set a unique ID for the table
        heading.textContent = difference;
        document.body.appendChild(heading);
    }
}

function createTable(sheetData, title, headerData) {
    const table = document.createElement('table');
    table.id = 'comparisonTable'; // Set a unique ID for the table
    const headerRow = table.insertRow();

    // Add header cells
    for (let i = 0; i < headerData[0].length; i++) {
        const cellData = headerData[0][i];
        const headerCell = headerRow.insertCell();
        headerCell.textContent = cellData;
    }

    // Add data rows for sheet1
    for (let i = 0; i < sheetData.length; i++) {
        const rowData = sheetData[i];
        const row = table.insertRow();
        for (let j = 0; j < rowData.length; j++) {
            const cellData = rowData[j];
            const cell = row.insertCell();
            cell.textContent = cellData;
        }
    }

    const heading = document.createElement('h2');
    heading.textContent = title;
    document.body.appendChild(heading);
    document.body.appendChild(table);
}

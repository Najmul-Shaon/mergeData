document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('exportButton').addEventListener('click', exportTableToExcel);

let jsonData = []; // Global variable to store the JSON data

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        displayTable(jsonData);
        document.getElementById('exportButton').disabled = false; // Enable export button
    };

    reader.readAsArrayBuffer(file);
}

function displayTable(data) {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');

    // Clear previous content
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        row.forEach(cell => {
            const cellElement = document.createElement(rowIndex === 0 ? 'th' : 'td');
            cellElement.textContent = cell || ''; // If cell is empty, display an empty string
            tr.appendChild(cellElement);
        });

        if (rowIndex === 0) {
            tableHead.appendChild(tr); // First row goes to table header
        } else {
            tableBody.appendChild(tr); // Remaining rows go to table body
        }
    });
}

function exportTableToExcel() {
    const worksheet = XLSX.utils.aoa_to_sheet(jsonData); // Convert JSON data back to worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Trigger the file download
    XLSX.writeFile(workbook, 'ExportedData.xlsx');
}

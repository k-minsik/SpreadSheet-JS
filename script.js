const spreadSheetContainer = document.querySelector("#spreadsheet-container");
const exportBtn = document.querySelector("#export-btn");
const cellStatus = document.querySelector("#cell-status");
const addRowBtn = document.getElementById('add-row-btn');
const removeRowBtn = document.getElementById('remove-row-btn');
const addColumnBtn = document.getElementById('add-column-btn');
const removeColumnBtn = document.getElementById('remove-column-btn');

const ROWS = 10;
const COLS = 10;
const alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split('');

class Cell {
    constructor(isHeader, data, row, column) {
        this.isHeader = isHeader;
        this.data = data;
        this.row = row;
        this.column = column;
    }
}

let spreadsheet = [];

function exportToCsv() {
    let csv = "";
    document.querySelectorAll('.cell-row').forEach((rowEl, rowIndex) => {
        if (rowIndex === 0) return;
        const row = [];
        rowEl.querySelectorAll('.cell').forEach(cellEl => {
            if (!cellEl.classList.contains('header')) {
                row.push(cellEl.value);
            }
        });
        csv += row.join(',') + "\r\n";
    });
    const csvBlob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const csvUrl = URL.createObjectURL(csvBlob);
    const a = document.createElement("a");
    a.href = csvUrl;
    a.download = 'spreadsheet.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

exportBtn.addEventListener('click', exportToCsv);

function initSpreadsheet() {
    spreadsheet = [];
    for (let i = 0; i < ROWS; i++) {
        const newRow = [];
        for (let j = 0; j < COLS; j++) {
            const isHeader = i === 0 || j === 0;
            const data = isHeader ? (j === 0 ? i : alphabets[j - 1]) : "";
            newRow.push(new Cell(isHeader, data, i, j));
        }
        spreadsheet.push(newRow);
    }
    drawSheet();
}

function createCellEl(cell) {
    const cellEl = document.createElement('input');
    cellEl.type = 'text';
    cellEl.className = `cell ${cell.isHeader ? 'header' : ''}`;
    cellEl.id = `cell_${cell.row}_${cell.column}`;
    cellEl.value = cell.data;
    cellEl.disabled = cell.isHeader;
    return cellEl;
}

function drawSheet() {
    spreadSheetContainer.innerHTML = '';
    for (let i = 0; i < spreadsheet.length; i++) {
        const rowContainerEl = document.createElement("div");
        rowContainerEl.className = "cell-row";
        for (let j = 0; j < spreadsheet[i].length; j++) {
            const cell = spreadsheet[i][j];
            rowContainerEl.appendChild(createCellEl(cell));
        }
        spreadSheetContainer.appendChild(rowContainerEl);
    }
}

function addRow() {
    const newRow = [];
    const newRowNum = spreadsheet.length;
    for (let i = 0; i < spreadsheet[0].length; i++) {
        const isHeader = i === 0;
        const cellData = isHeader ? newRowNum : '';
        newRow.push(new Cell(isHeader, cellData, newRowNum, i));
    }
    spreadsheet.push(newRow);
    drawSheet();
}

function removeRow() {
    if (spreadsheet.length > 1) {
        spreadsheet.pop();
        drawSheet();
    }
}

function addColumn() {
    const newColIndex = spreadsheet[0].length;
    const alphabetIndex = newColIndex >= alphabets.length ? newColIndex : alphabets[newColIndex - 1];
    spreadsheet.forEach((row, rowIndex) => {
        const isHeader = rowIndex === 0;
        const cellData = isHeader ? alphabetIndex : '';
        row.push(new Cell(isHeader, cellData, rowIndex, newColIndex));
    });
    drawSheet();
}

function removeColumn() {
    if (spreadsheet[0].length > 1) {
        spreadsheet.forEach(row => row.pop());
        drawSheet();
    }
}

addRowBtn.addEventListener('click', addRow);
removeRowBtn.addEventListener('click', removeRow);
addColumnBtn.addEventListener('click', addColumn);
removeColumnBtn.addEventListener('click', removeColumn);

spreadSheetContainer.addEventListener('click', function(event) {
    const cellEl =
    event.target.closest('.cell');
    if (!cellEl) return;

    clearHeaderActiveStates();
    
    const row = cellEl.parentElement;
    const cellIndex = Array.from(row.children).indexOf(cellEl);
    
    if (cellIndex > 0) {
        const columnHeaderEl = spreadSheetContainer.querySelector(`#cell_0_${cellIndex}`);
        columnHeaderEl.classList.add('active');
    }
    
    const rowIndex = Array.from(spreadSheetContainer.children).indexOf(row);
    if (rowIndex > 0) {
        const rowHeaderEl = spreadSheetContainer.querySelector(`#cell_${rowIndex}_0`);
        rowHeaderEl.classList.add('active');
    }

    const columnName = alphabets[cellIndex - 1] || "";
    cellStatus.textContent = `${columnName}${rowIndex}`;
});

spreadSheetContainer.addEventListener('input', function(event) {
    const cellEl = event.target.closest('.cell');
    if (cellEl && !cellEl.classList.contains('header')) {
        const rowIndex = cellEl.id.split('_')[1];
        const columnIndex = cellEl.id.split('_')[2];
        spreadsheet[rowIndex][columnIndex].data = cellEl.value;
    }
});

function clearHeaderActiveStates() {
    document.querySelectorAll('.header.active').forEach(header => {
        header.classList.remove('active');
    });
}

initSpreadsheet();

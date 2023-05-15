// Constants
const ROWS = 100;
const COLUMNS = 100;
const SUPPORTED_FUNCTIONS = {
  SUM: sumRange,
}; // Add more supported functions to increase functionality

// Object to store cell data
const cells = {};
const columnHeaders = getColumnHeaders(COLUMNS);

// Generates an array of column headers based on the count provided.
// Outputs an array such as  A, B, C, ..., Z, AA, AB, AC, ...
function getColumnHeaders(count) {
  const columnHeaders = [];
  const charCodeOffset = 65; // ASCII code for 'A'
  let columnCount = 1;

  while (columnCount <= count) {
    let header = "";
    let quotient = columnCount - 1;

    while (quotient >= 0) {
      const remainder = quotient % 26;
      header = String.fromCharCode(charCodeOffset + remainder) + header;
      quotient = Math.floor(quotient / 26) - 1;
    }

    columnHeaders.push(header);
    columnCount++;
  }

  return columnHeaders;
}

// Parse cell ID into column and row
// For eg, A10 will be parsed into [A, 10]
function parseCellId(cellId) {
  const col = cellId.match(/[A-Z]+/)[0];
  const row = parseInt(cellId.match(/\d+/)[0]);
  return [col, row];
}

// Generate the spreadsheet
function generateSpreadsheet() {
  const table = document.getElementById("spreadsheet");
  const tbody = document.createElement("tbody");

  // Create header row
  const header = document.createElement("tr");
  header.appendChild(document.createElement("th")); // Empty column for row labels

  // Add column labels
  columnHeaders.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    header.appendChild(th);
  });

  tbody.appendChild(header);

  // Create rows and columns
  for (let row = 1; row <= ROWS; row++) {
    const tr = document.createElement("tr");
    const th = document.createElement("th");
    th.textContent = row;
    tr.appendChild(th);

    for (let col = 0; col < COLUMNS; col++) {
      const div = createEditableDiv(columnHeaders[col] + row);
      const td = document.createElement("td");
      td.appendChild(div);
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
}

// Create an editable div element
function createEditableDiv(id) {
  const div = document.createElement("div");
  div.id = id;
  div.contentEditable = true;
  div.addEventListener("blur", evaluateFormula);
  return div;
}

// Evaluate the formula entered in a cell
function evaluateFormula(event) {
  const cell = event.target;
  const cellId = cell.id;
  const formula = cell.innerText.trim();

  if (formula.startsWith("=")) {
    const updatedFormula = formula.toUpperCase();
    const expression = updatedFormula.substring(1); // Remove the leading "="
    const evaluatedValue = evaluateExpression(expression);
    cells[cellId] = { updatedFormula, value: evaluatedValue };
    cell.innerText = evaluatedValue;
  } else {
    cells[cellId] = { formula: "", value: formula }; // Store non-formula values directly
  }

  updateDependentCells(cellId);
}

// Evaluate the expression in a cell
// Expression can either be a function or direct computation
// Supported function format: function_name(Range). 
//    function_name: should be a property of SUPPORTED_FUNCTIONS
//    Range: should be of format (CellId1: CellId2). Eg (A1: C1)
// Only basic computation is provided, without support of paranthesis
function evaluateExpression(expression, col, row) {
  const functionRegex = /([A-Z]+)\([A-Z]+\d+:[A-Z]+\d+\)/;
  if (functionRegex.test(expression)) {
    // Handle functions such as SUM, AVERAGE etc.
    const functionMatch = expression.match(functionRegex)[0];
    const functionName = functionMatch.split("(")[0].replace("=", "");
    if (SUPPORTED_FUNCTIONS.hasOwnProperty(functionName)) {
      expression = expression.replace(
        functionMatch,
        evaluateFunction(functionMatch, functionName)
      );
    } else {
      console.error("Error evaluating expression: Invalid function name");
      return "N/A"; // Return empty string on error
    }
  } else {
    // Handle basic computations such as + - * /
    const cellReferenceRegex = /[A-Z]+\d+/g;
    const cellReferences = expression.match(cellReferenceRegex) || [];
    cellReferences.forEach((cellReference) => {
      const [refCol, refRow] = parseCellId(cellReference);
      const cellId = refCol + refRow;
      const cellValue = cells[cellId] || 0; // Treat empty cells as 0
      expression = expression.replace(cellReference, cellValue.value);
    });
  }

  try {
    const result = eval(expression);
    return isNaN(result) ? "" : result; // Return empty string if the result is not a valid number
  } catch (error) {
    console.error("Error evaluating expression:", error);
    return "N/A"; // Return empty string on error
  }
}

// Evaluate a function expression
function evaluateFunction(functionExpression, functionName) {
  const rangeRegex = /[A-Z]+\d+:[A-Z]+\d+/;
  const range = functionExpression.match(rangeRegex)[0];
  const [startCell, endCell] = range.split(":");
  const [startCol, startRow] = parseCellId(startCell);
  const [endCol, endRow] = parseCellId(endCell);

  if (SUPPORTED_FUNCTIONS.hasOwnProperty(functionName)) {
    const getFunction = SUPPORTED_FUNCTIONS[functionName];
    return getFunction(startCol, startRow, endCol, endRow);
  }

  console.error("Error evaluating function:", functionName);
  return ""; // Return empty string on error
}

// Update formulas that depend on updated cell
function updateDependentCells(updatedCellId) {
  Object.keys(cells).forEach((cellId) => {
    const { formula, value } = cells[cellId];
    if (formula.includes(updatedCellId)) {
      const [col, row] = parseCellId(cellId);
      const evaluatedValue = evaluateExpression(formula.substring(1), col, row);
      cells[cellId] = { formula, value: evaluatedValue };
      const cellElement = document.getElementById(cellId);
      cellElement.innerText = evaluatedValue;
    }
  });
}

// Calculate the sum of a range of cells
function sumRange(startCol, startRow, endCol, endRow) {
  let sum = 0;
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
      const cellId = String.fromCharCode(col) + row;
      const cellValue = cells[cellId] || 0;
      sum += parseFloat(cellValue.value);
    }
  }
  return sum;
}

// Refresh the spreadsheet by clearing div contents
function refreshSheet() {
  const cells = Array.from(document.querySelectorAll("#spreadsheet td div"));
  cells.forEach((cell) => (cell.innerText = ""));
}

// Make refresh button work
function setRefresh() {
  const refreshBtn = document.getElementById("refreshBtn");
  refreshBtn.addEventListener("click", refreshSheet);
  refreshBtn.disabled = false;
}

// JavaScript object to store cell values
document.addEventListener("DOMContentLoaded", generateSpreadsheet);
document.addEventListener("DOMContentLoaded", setRefresh);

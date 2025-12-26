import ExcelJS from 'exceljs';
import path from 'path';

export const getPreviewData = async (filePath) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.worksheets[0];
    
    const data = [];
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        // row.values is 1-based, index 0 is undefined/null usually
        // We want to return a clean array of arrays or objects
        const rowValues = JSON.parse(JSON.stringify(row.values));
        // Remove the first element if it's null/undefined (ExcelJS quirk)
        if (Array.isArray(rowValues) && (rowValues[0] === null || rowValues[0] === undefined)) {
            rowValues.shift();
        }
        // ExcelJS might return sparse arrays, fill with null/empty string
        for(let i=0; i<rowValues.length; i++) {
             if(rowValues[i] === undefined) rowValues[i] = null;
        }
        data.push(rowValues);
    });
    
    return data;
};

export const processExcelAction = async (filePath, actionData) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0]; // Assume first sheet

  const { action, params } = actionData;

  switch (action) {
    case 'ADD_COLUMN':
      await addColumn(worksheet, params);
      break;
    case 'HIGHLIGHT_ROWS':
      await highlightRows(worksheet, params);
      break;
    case 'SORT_DATA':
      await sortData(worksheet, params);
      break;
    default:
      // For now ignore unknown actions or throw
      console.warn(`Unsupported action: ${action}`);
  }

  const parsed = path.parse(filePath);
  const newFilePath = path.join(parsed.dir, `${parsed.name}_${Date.now()}${parsed.ext}`);

  await workbook.xlsx.writeFile(newFilePath);
  return newFilePath;
};

const getColumnIndex = (worksheet, name) => {
    const headerRow = worksheet.getRow(1);
    let colIndex = -1;
    headerRow.eachCell((cell, colNumber) => {
        if (cell.value && cell.value.toString().toLowerCase() === name.trim().toLowerCase()) {
            colIndex = colNumber;
        }
    });
    return colIndex;
};

const colLetter = (col) => {
    let temp, letter = '';
    while (col > 0) {
      temp = (col - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
};

const addColumn = async (worksheet, { columnName, formula }) => {
    const headerRow = worksheet.getRow(1);
    const nextCol = headerRow.cellCount + 1;
    worksheet.getCell(1, nextCol).value = columnName;

    const headers = {};
    headerRow.eachCell((cell, colNumber) => {
        headers[cell.value.toString().toLowerCase()] = colNumber;
    });

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        
        let parsedFormula = formula;
        Object.keys(headers).forEach(header => {
             const regex = new RegExp(`\\b${header}\\b`, 'gi');
             const colIdx = headers[header];
             parsedFormula = parsedFormula.replace(regex, `${colLetter(colIdx)}${rowNumber}`);
        });

        worksheet.getCell(rowNumber, nextCol).value = { formula: parsedFormula };
    });
};

const highlightRows = async (worksheet, { condition, color }) => {
    const operators = ['>=', '<=', '!=', '>', '<', '='];
    let operator = operators.find(op => condition.includes(op));
    if (!operator) return;

    const [colName, value] = condition.split(operator).map(s => s.trim());
    const colIndex = getColumnIndex(worksheet, colName);
    
    if (colIndex === -1) return;

    const threshold = parseFloat(value);

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const cellValue = row.getCell(colIndex).value;
        let match = false;
        
        const numVal = parseFloat(cellValue);
        
        if (!isNaN(numVal)) {
            switch(operator) {
                case '>': match = numVal > threshold; break;
                case '<': match = numVal < threshold; break;
                case '>=': match = numVal >= threshold; break;
                case '<=': match = numVal <= threshold; break;
                case '=': match = numVal === threshold; break;
                case '!=': match = numVal !== threshold; break;
            }
        }
        
        if (match) {
            row.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: color || 'FFFF00' }
                };
            });
        }
    });
};

const sortData = async (worksheet, { column, order }) => {
    const colIndex = getColumnIndex(worksheet, column);
    if (colIndex === -1) return;

    // Get all data rows
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        rows.push({
            values: row.values, // Note: row.values is 1-based array usually, index 0 is empty? Check docs.
            sortValue: row.getCell(colIndex).value
        });
    });

    // Sort
    rows.sort((a, b) => {
        const valA = a.sortValue;
        const valB = b.sortValue;
        if (valA > valB) return order === 'desc' ? -1 : 1;
        if (valA < valB) return order === 'desc' ? 1 : -1;
        return 0;
    });

    // Write back
    // Warning: row.values setter might need exact array structure.
    // ExcelJS `row.values` includes index 0 as undefined? 
    // "The values property returns an array of values where the index corresponds to the column number."
    // So [undefined, col1, col2, ...]
    
    rows.forEach((r, i) => {
        const targetRow = worksheet.getRow(i + 2);
        targetRow.values = r.values;
    });
};

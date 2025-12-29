import ExcelJS from 'exceljs';
import path from 'path';

export const getPreviewData = async (filePath) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const sheets = [];

    workbook.eachSheet((worksheet, sheetId) => {
        const sheetData = [];
        // Ensure we capture all rows, even if they are empty, up to the last used row
        const rowCount = worksheet.rowCount;
        
        // Iterate up to rowCount manually to ensure we don't miss anything
        // eachRow with includeEmpty might still be sparse
        for (let i = 1; i <= rowCount; i++) {
             const row = worksheet.getRow(i);
             const rowValues = JSON.parse(JSON.stringify(row.values));
             
             // Remove the first element if it's null/undefined (ExcelJS quirk: index 0 is reserved)
             if (Array.isArray(rowValues) && (rowValues[0] === null || rowValues[0] === undefined)) {
                 rowValues.shift();
             } else if (!Array.isArray(rowValues)) {
                 // If rowValues is not an array (e.g. object), handle it or default to empty
                 // For empty rows, row.values might be undefined or just { ... }
                 // If it's an object (Rich Text?), we might need more complex parsing, but for now assume standard values
             }
             
             // Ensure sparse arrays are filled with nulls
             // Note: rowValues might be shorter than the max column count.
             // We should ideally pad it to the max column count of the sheet, but the frontend handles varying lengths.
             if (Array.isArray(rowValues)) {
                 for(let j=0; j<rowValues.length; j++) {
                     if(rowValues[j] === undefined) rowValues[j] = null;
                 }
                 sheetData.push(rowValues);
             } else {
                 sheetData.push([]); // Empty row
             }
        }
        
        sheets.push({
            name: worksheet.name,
            data: sheetData
        });
    });
    
    return { sheets };
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

/**
 * Get windowed/sliced data from an Excel sheet
 * This function loads the full workbook but returns only the requested window
 * to enable virtualized rendering on the frontend.
 * 
 * @param {string} filePath - Path to the Excel file
 * @param {number} sheetIndex - Index of the sheet (0-based)
 * @param {number} rowStart - Starting row index (1-based, Excel convention)
 * @param {number} rowEnd - Ending row index (1-based, inclusive)
 * @param {number} colStart - Starting column index (1-based, Excel convention)
 * @param {number} colEnd - Ending column index (1-based, inclusive)
 * @returns {Promise<Object>} Windowed data with metadata
 */
/**
 * Get workbook metadata (all sheets with dimensions)
 * Returns sheet information without loading cell data
 * 
 * @param {string} filePath - Path to the Excel file
 * @returns {Promise<Object>} Workbook metadata with sheets array
 */
export const getWorkbookMetadata = async (filePath) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const sheets = [];
    workbook.eachSheet((worksheet, sheetId) => {
        sheets.push({
            sheetId: sheetId - 1, // Convert to 0-based for frontend
            name: worksheet.name,
            totalRows: worksheet.rowCount || 0,
            totalCols: worksheet.columnCount || 0
        });
    });
    
    return { sheets };
};

/**
 * Update a single cell in the workbook
 * 
 * @param {string} filePath - Path to the Excel file
 * @param {number} sheetIndex - Sheet index (0-based)
 * @param {number} row - Row number (1-based, Excel convention)
 * @param {number} col - Column number (1-based, Excel convention)
 * @param {any} value - New cell value
 * @returns {Promise<void>}
 */
export const updateCell = async (filePath, sheetIndex, row, col, value) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Validate sheet index
    if (sheetIndex < 0 || sheetIndex >= workbook.worksheets.length) {
        throw new Error(`Sheet index ${sheetIndex} is out of range. Total sheets: ${workbook.worksheets.length}`);
    }
    
    const worksheet = workbook.worksheets[sheetIndex];
    const cell = worksheet.getCell(row, col);
    
    // Set the cell value
    cell.value = value;
    
    // Save the workbook back to file
    await workbook.xlsx.writeFile(filePath);
};

export const getWindowedSheetData = async (filePath, sheetIndex, rowStart, rowEnd, colStart, colEnd) => {
    // Load the full workbook (required by ExcelJS)
    // This is done once per request - in production, you might want to cache workbooks
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Validate sheet index
    if (sheetIndex < 0 || sheetIndex >= workbook.worksheets.length) {
        throw new Error(`Sheet index ${sheetIndex} is out of range. Total sheets: ${workbook.worksheets.length}`);
    }
    
    const worksheet = workbook.worksheets[sheetIndex];
    const sheetName = worksheet.name;
    
    // Get total dimensions of the sheet
    // ExcelJS rowCount and columnCount give us the actual used dimensions
    const totalRows = worksheet.rowCount || 0;
    const totalColumns = worksheet.columnCount || 0;
    
    // Clamp the requested window to actual sheet bounds
    // ExcelJS uses 1-based indexing for rows and columns
    const clampedRowStart = Math.max(1, Math.min(rowStart, totalRows + 1));
    const clampedRowEnd = Math.max(1, Math.min(rowEnd, totalRows));
    const clampedColStart = Math.max(1, Math.min(colStart, totalColumns + 1));
    const clampedColEnd = Math.max(1, Math.min(colEnd, totalColumns));
    
    // Extract only the requested window of data
    // This prevents sending the entire sheet to the frontend
    const windowData = [];
    
    // Iterate only through the requested rows (1-based in ExcelJS)
    for (let rowNum = clampedRowStart; rowNum <= clampedRowEnd; rowNum++) {
        const row = worksheet.getRow(rowNum);
        const rowValues = [];
        
        // Extract only the requested columns (1-based in ExcelJS)
        for (let colNum = clampedColStart; colNum <= clampedColEnd; colNum++) {
            const cell = row.getCell(colNum);
            let cellValue = cell.value;
            
            // Handle different cell value types
            // ExcelJS can return various types: string, number, Date, formula result, etc.
            if (cellValue === null || cellValue === undefined) {
                cellValue = null;
            } else if (cellValue instanceof Date) {
                // Convert dates to ISO string for JSON serialization
                cellValue = cellValue.toISOString();
            } else if (typeof cellValue === 'object' && cellValue !== null) {
                // For complex objects (formulas, rich text), extract the text representation
                // In production, you might want to handle formulas differently
                if (cellValue.text) {
                    cellValue = cellValue.text;
                } else if (cellValue.result) {
                    // Formula result
                    cellValue = cellValue.result;
                } else {
                    // Fallback: stringify complex objects
                    cellValue = JSON.stringify(cellValue);
                }
            }
            
            rowValues.push(cellValue);
        }
        
        windowData.push(rowValues);
    }
    
    // Return windowed data with metadata
    // The frontend uses this to:
    // 1. Render only the visible cells
    // 2. Know the total dimensions for virtual scrolling
    // 3. Calculate which window to fetch next on scroll
    return {
        data: windowData,
        meta: {
            totalRows,
            totalColumns,
            sheetName,
            // Return the actual window bounds that were returned (useful for debugging)
            window: {
                rowStart: clampedRowStart,
                rowEnd: clampedRowEnd,
                colStart: clampedColStart,
                colEnd: clampedColEnd
            }
        }
    };
};

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



const getColumnIndex = (worksheet, name, scanDepth = 20) => {
    let colIndex = -1;
    let foundAtRow = -1;

    // Scan first N rows to find the header/key
    worksheet.eachRow((row, rowNumber) => {
        if (colIndex !== -1 || rowNumber > scanDepth) return;
        
        row.eachCell((cell, colNumber) => {
            if (colIndex !== -1) return;
            if (cell.value && cell.value.toString().trim().toLowerCase() === name.trim().toLowerCase()) {
                colIndex = colNumber;
                foundAtRow = rowNumber;
            }
        });
    });
    
    return { colIndex, foundAtRow };
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
    const { colIndex } = getColumnIndex(worksheet, colName);
    
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
    const { colIndex, foundAtRow } = getColumnIndex(worksheet, column);
    if (colIndex === -1) {
        throw new Error(`Column "${column}" not found.`);
    }

    console.log(`[SORT_DATA] Sorting by "${column}" (Index: ${colIndex}, Header Row: ${foundAtRow})`);

    // Get all data rows below the header
    const rows = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= foundAtRow) return; // Skip header and above
        
        rows.push({
            originalRowNumber: rowNumber,
            values: row.values, 
            sortValue: row.getCell(colIndex).value
        });
    });

    if (rows.length === 0) {
        console.warn('[SORT_DATA] No data rows found to sort.');
        return;
    }

    // Sort
    rows.sort((a, b) => {
        const valA = a.sortValue;
        const valB = b.sortValue;

        // Handle nulls/undefineds - push to bottom
        if (valA == null && valB == null) return 0;
        if (valA == null) return 1; 
        if (valB == null) return -1;

        // Try to convert to numbers for comparison
        // We use Number() instead of parseFloat() to avoid "12abc" being treated as 12.
        const numA = Number(valA);
        const numB = Number(valB);
        
        // Check if valid numbers (and not empty strings which Number() converts to 0)
        const isNumA = !isNaN(numA) && valA.toString().trim() !== '';
        const isNumB = !isNaN(numB) && valB.toString().trim() !== '';

        if (isNumA && isNumB) {
            return order === 'desc' ? numB - numA : numA - numB;
        }

        // Fallback to string sorting
        const strA = valA.toString().toLowerCase();
        const strB = valB.toString().toLowerCase();
        
        if (strA > strB) return order === 'desc' ? -1 : 1;
        if (strA < strB) return order === 'desc' ? 1 : -1;
        return 0;
    });

    // Write back to the same row numbers we read from (preserving gaps if any)
    const targetRowNumbers = rows.map(r => r.originalRowNumber).sort((a, b) => a - b);
    
    rows.forEach((sortedRow, i) => {
        const targetRowIdx = targetRowNumbers[i];
        const targetRow = worksheet.getRow(targetRowIdx);
        targetRow.values = sortedRow.values;
    });
    
    console.log(`[SORT_DATA] Sorted ${rows.length} rows.`);
};

const updateKeyValue = async (worksheet, { keyColumn, keyValue, valueColumn, newValue }) => {
    console.log(`[UPDATE_KEY_VALUE] Started. KeyCol: ${keyColumn}, KeyVal: ${keyValue}, ValCol: ${valueColumn}, NewVal: ${newValue}`);

    // 1. Find Key Column
    const { colIndex: keyColIndex, foundAtRow: keyHeaderRow } = getColumnIndex(worksheet, keyColumn);
    if (keyColIndex === -1) {
        throw new Error(`Key Column "${keyColumn}" not found.`);
    }

    // 2. Find Value Column
    let valColIndex = -1;
    if (valueColumn) {
        const { colIndex } = getColumnIndex(worksheet, valueColumn);
        valColIndex = colIndex;
    }

    // Fallback: If Value Column not specified or not found, assume it's the NEXT column (standard Key-Value pattern)
    if (valColIndex === -1) {
        valColIndex = keyColIndex + 1;
        console.log(`[UPDATE_KEY_VALUE] Value Column not explicitly found. Defaulting to Next Column (Index: ${valColIndex})`);
    }

    let rowsUpdated = 0;

    worksheet.eachRow((row, rowNumber) => {
        // We DO NOT skip the header row for UPDATE_KEY_VALUE
        // Reason: In Key-Value tables, the "header" (Key Label) IS the row we want to target.
        // If it's a standard table, the header value usually won't match the target value anyway.

        const keyCell = row.getCell(keyColIndex);
        const keyVal = keyCell.value;

        // Loose matching for KEY
        const isKeyMatch = keyVal == keyValue || 
                       (keyVal && keyValue && keyVal.toString().trim().toLowerCase() === keyValue.toString().trim().toLowerCase());

        // Loose matching for VALUE (Fallback: AI might send the old value as the key)
        const valCell = row.getCell(valColIndex);
        const valVal = valCell.value;
        const isValueMatch = valVal == keyValue || 
                       (valVal && keyValue && valVal.toString().trim().toLowerCase() === keyValue.toString().trim().toLowerCase());

        if (isKeyMatch || isValueMatch) {
            const targetCell = row.getCell(valColIndex);
            console.log(`[UPDATE_KEY_VALUE] Updating Cell [${rowNumber}, ${valColIndex}]: ${targetCell.value} -> ${newValue}`);
            targetCell.value = newValue;
            rowsUpdated++;
        }
    });

    if (rowsUpdated === 0) {
        throw new Error(`No rows updated. Key "${keyValue}" not found in column "${keyColumn}".`);
    }

    console.log(`[UPDATE_KEY_VALUE] Finished. Rows updated: ${rowsUpdated}`);
    return rowsUpdated;
};

const setCell = async (worksheet, { cell, value }) => {
    console.log(`[SET_CELL] Setting ${cell} to ${value}`);
    const targetCell = worksheet.getCell(cell);
    targetCell.value = value;
    console.log(`[SET_CELL] Updated ${cell}`);
};

export const findAndReplace = async (worksheet, { findValue, replaceValue, column }) => {
    console.log(`[FIND_AND_REPLACE] Started. Find: "${findValue}", Replace: "${replaceValue}", Column: "${column || 'ALL'}"`);

    let targetColIndex = -1;
    if (column) {
        const { colIndex } = getColumnIndex(worksheet, column);
        if (colIndex !== -1) {
            targetColIndex = colIndex;
            console.log(`[FIND_AND_REPLACE] Restricted to column "${column}" (Index: ${targetColIndex})`);
        } else {
             // If user specified a column but we can't find it as a header, 
             // we should probably warn or try to treat 'column' as a loose hint?
             // For strictness, let's error if specific column requested but missing.
             // BUT, if the user said "Change entc to 12333", 'entc' might be the value, not the column.
             // So if column not found, we might want to fallback to global search?
             // Let's stick to strict: if column provided, it must exist.
            throw new Error(`Column "${column}" not found.`);
        }
    }

    let rowsUpdated = 0;

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            // If column restriction is active, skip other columns
            if (targetColIndex !== -1 && colNumber !== targetColIndex) return;

            const cellValue = cell.value;
            // Loose matching
            if (cellValue == findValue || 
               (cellValue && findValue && cellValue.toString().trim().toLowerCase() === findValue.toString().trim().toLowerCase())) {
                
                console.log(`[FIND_AND_REPLACE] Updating Cell [${rowNumber}, ${colNumber}]: ${cellValue} -> ${replaceValue}`);
                cell.value = replaceValue;
                rowsUpdated++;
            }
        });
    });

    if (rowsUpdated === 0) {
        throw new Error(`No occurrences of "${findValue}" found${column ? ` in column "${column}"` : ''}.`);
    }

    return rowsUpdated;
};

const updateRowValues = async (worksheet, { filterColumn, filterValue, operation, value, targetColumn }) => {
    console.log(`[UPDATE_ROW_VALUES] Started. Filter: ${filterColumn}=${filterValue}, Op: ${operation}, Val: ${value}, Target: ${targetColumn}`);
    
    // 1. Find the filter column
    let searchColIndex = -1;
    let headerRowIndex = 1;
    
    if (filterColumn) {
        const { colIndex, foundAtRow } = getColumnIndex(worksheet, filterColumn);
        if (colIndex !== -1) {
            searchColIndex = colIndex;
            headerRowIndex = foundAtRow;
            console.log(`[UPDATE_ROW_VALUES] Column "${filterColumn}" found at index: ${colIndex} (Row ${foundAtRow})`);
        } else {
            throw new Error(`Filter Column "${filterColumn}" not found.`);
        }
    } else {
        throw new Error(`Filter Column is required.`);
    }

    // 2. Determine Target Column
    let targetColIndex = -1;
    if (targetColumn) {
        const { colIndex } = getColumnIndex(worksheet, targetColumn);
        if (colIndex !== -1) {
            targetColIndex = colIndex;
            console.log(`[UPDATE_ROW_VALUES] Target Column "${targetColumn}" found at index: ${targetColIndex}`);
        } else {
            throw new Error(`Target Column "${targetColumn}" not found.`);
        }
    } else if (operation === 'SET') {
        throw new Error(`Target Column is required for SET operation.`);
    }

    let rowsUpdated = 0;

    worksheet.eachRow((row, rowNumber) => {
        // Skip header row if we found one, otherwise process all
        if (rowNumber === headerRowIndex && headerRowIndex !== -1) return; 

        // Check if this row matches the criteria
        const keyCell = row.getCell(searchColIndex);
        const keyVal = keyCell.value;

        // Loose matching (string vs number, case insensitive)
        const isMatch = keyVal == filterValue || 
                       (keyVal && filterValue && keyVal.toString().trim().toLowerCase() === filterValue.toString().trim().toLowerCase());

        if (isMatch) {
            console.log(`[UPDATE_ROW_VALUES] Match found at row ${rowNumber}. Key: ${keyVal}`);
            
            // Define update logic
            const updateCell = (cell) => {
                const currentVal = cell.value;
                let newVal = currentVal;
                
                if (operation === 'SET') {
                    newVal = value;
                } else {
                    // Arithmetic operations (numbers only)
                    const numCurrent = parseFloat(currentVal);
                    // Ensure value is treated as number for arithmetic
                    const numValue = parseFloat(value);
                    
                    if (!isNaN(numCurrent) && !isNaN(numValue)) {
                        switch (operation) {
                            case '+': newVal = numCurrent + numValue; break;
                            case '-': newVal = numCurrent - numValue; break;
                            case '*': newVal = numCurrent * numValue; break;
                            case '/': if (numValue !== 0) newVal = numCurrent / numValue; break;
                        }
                    } else {
                        return; // Skip non-numeric cells for arithmetic
                    }
                }
                
                if (newVal !== currentVal) {
                    console.log(`[UPDATE_ROW_VALUES] Updating Cell [${rowNumber}, ${cell.col}]: ${currentVal} -> ${newVal}`);
                    cell.value = newVal;
                }
            };

            if (targetColIndex !== -1) {
                // Update specific column
                const targetCell = row.getCell(targetColIndex);
                updateCell(targetCell);
            } else {
                // Update ALL cells in the row (ONLY for arithmetic)
                // We deliberately exclude SET here to prevent accidental row wipes
                row.eachCell((cell, colNum) => {
                    if (colNum !== searchColIndex) {
                        updateCell(cell);
                    }
                });
            }
            rowsUpdated++;
        }
    });
    
    if (rowsUpdated === 0) {
        throw new Error(`No rows updated. Filter value "${filterValue}" not found in column "${filterColumn}".`);
    }

    console.log(`[UPDATE_ROW_VALUES] Finished. Rows updated: ${rowsUpdated}`);
};

export const processExcelAction = async (filePath, actionData) => {
  console.log('[processExcelAction] Loading workbook:', filePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  // 2. Verify Worksheet Resolution
  const worksheet = workbook.worksheets[0]; 
  if (!worksheet) {
      throw new Error("Worksheet not found in the workbook.");
  }
  console.log(`[processExcelAction] Using sheet: ${worksheet.name}`);

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
    case 'ADD_VALUE_TO_ROW': // Legacy support
    case 'UPDATE_ROW_VALUES':
      await updateRowValues(worksheet, params);
      break;
    case 'UPDATE_KEY_VALUE':
      await updateKeyValue(worksheet, params);
      break;
    case 'SET_CELL':
      await setCell(worksheet, params);
      break;
    case 'FIND_AND_REPLACE':
      await findAndReplace(worksheet, params);
      break;
    default:
      console.warn(`Unsupported action: ${action}`);
      throw new Error(`Unsupported action: ${action}`);
  }

  const parsed = path.parse(filePath);
  const newFilePath = path.join(parsed.dir, `${parsed.name}_${Date.now()}${parsed.ext}`);

  // 4. Save Workbook (Critical)
  await workbook.xlsx.writeFile(newFilePath);
  
  // 5. Confirm Persistence
  console.log("Workbook saved successfully");
  console.log(`[processExcelAction] Saved to: ${newFilePath}`);
  
  return newFilePath;
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
 * CRITICAL RULES:
 * - Only include visible sheets
 * - Use sheet name as ID (NOT index)
 * - Preserve Excel tab order
 * - Use dimensions for accurate row/col counts
 * 
 * @param {string} filePath - Path to the Excel file
 * @returns {Promise<Object>} Workbook metadata with sheets array
 */
export const getWorkbookMetadata = async (filePath) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Filter only visible sheets and preserve Excel tab order
    const sheets = workbook.worksheets
        .filter(ws => ws.state === 'visible' || ws.state === undefined) // Include visible or undefined (default is visible)
        .map(worksheet => {
            const MAX_VIRTUAL_ROWS = 1000;
            const MAX_VIRTUAL_COLS = 100;
            return {
                sheetId: worksheet.name,  // Use name as ID, NOT index
                name: worksheet.name,
                totalRows: MAX_VIRTUAL_ROWS,
                totalCols: MAX_VIRTUAL_COLS
            };
        });
    
    return { sheets };
};

/**
 * Update a single cell in the workbook
 * 
 * @param {string} filePath - Path to the Excel file
 * @param {string} sheetId - Sheet name (NOT index)
 * @param {number} row - Row number (1-based, Excel convention)
 * @param {number} col - Column number (1-based, Excel convention)
 * @param {any} value - New cell value
 * @returns {Promise<void>}
 */
export const updateCell = async (filePath, sheetId, row, col, value) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  // Get worksheet by name (NOT by index)
  const worksheet = workbook.getWorksheet(sheetId);
  
  if (!worksheet) {
    throw new Error(`Sheet "${sheetId}" not found in workbook. Available sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
  }
  
  const cell = worksheet.getCell(row, col);
  
  // Set the cell value
  cell.value = value;
  
  // Save the workbook back to file
  await workbook.xlsx.writeFile(filePath);
};

export const addSheet = async (filePath, name) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const existing = workbook.getWorksheet(name);
  if (existing) {
    throw new Error(`Sheet "${name}" already exists`);
  }
  workbook.addWorksheet(name);
  await workbook.xlsx.writeFile(filePath);
  return { name };
};

export const renameSheet = async (filePath, sheetId, newName) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(sheetId);
  if (!worksheet) {
    throw new Error(`Sheet "${sheetId}" not found`);
  }
  if (workbook.getWorksheet(newName)) {
    throw new Error(`Sheet "${newName}" already exists`);
  }
  worksheet.name = newName;
  await workbook.xlsx.writeFile(filePath);
  return { oldName: sheetId, newName };
};

export const deleteSheet = async (filePath, sheetId) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(sheetId);
  if (!worksheet) {
    throw new Error(`Sheet "${sheetId}" not found`);
  }
  workbook.removeWorksheet(worksheet.id);
  await workbook.xlsx.writeFile(filePath);
  return { deleted: sheetId };
};

/**
 * Get windowed/sliced data from an Excel sheet
 * This function loads the full workbook but returns only the requested window
 * to enable virtualized rendering on the frontend.
 * 
 * CRITICAL REQUIREMENTS:
 * - Use sheet name (NOT index) for sheet selection
 * - Use dimensions for accurate row/col counts
 * - Preserve empty cells with index-based access (maintains column alignment)
 * - Never use eachCell (skips empty cells and breaks alignment)
 * 
 * @param {string} filePath - Path to the Excel file
 * @param {string} sheetId - Name of the sheet (NOT index)
 * @param {number} rowStart - Starting row index (1-based, Excel convention)
 * @param {number} rowEnd - Ending row index (1-based, inclusive)
 * @param {number} colStart - Starting column index (1-based, Excel convention)
 * @param {number} colEnd - Ending column index (1-based, inclusive)
 * @returns {Promise<Object>} Windowed data with metadata
 */
export const getWindowedSheetData = async (filePath, sheetId, rowStart, rowEnd, colStart, colEnd) => {
    // Load the full workbook (required by ExcelJS)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Get worksheet by name (NOT by index)
    const worksheet = workbook.getWorksheet(sheetId);
    
    if (!worksheet) {
        throw new Error(`Sheet "${sheetId}" not found in workbook. Available sheets: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
    }
    
    const sheetName = worksheet.name;
    
    // CRITICAL: Use dimensions for accurate counts (handles sparse data correctly)
    const dimensions = worksheet.dimensions;
    const actualRows = dimensions ? dimensions.bottom : (worksheet.rowCount || 0);
    const actualCols = dimensions ? dimensions.right : (worksheet.columnCount || 0);
    
    // Cap totals strictly for performance
    const MAX_VIRTUAL_ROWS = 1000;
    const MAX_VIRTUAL_COLS = 100;
    
    // Clamp the requested window to VIRTUAL bounds, not actual bounds
    // ExcelJS uses 1-based indexing for rows and columns
    const clampedRowStart = Math.max(1, Math.min(rowStart, MAX_VIRTUAL_ROWS));
    const clampedRowEnd = Math.max(1, Math.min(rowEnd, MAX_VIRTUAL_ROWS)); 
    const clampedColStart = 1; // No column windowing
    const clampedColEnd = MAX_VIRTUAL_COLS;
    
    // CRITICAL: Extract window data using index-based access
    // This preserves empty cells and maintains column alignment
    const windowData = [];
    
    // Iterate through requested rows (1-based in ExcelJS)
    for (let rowNum = clampedRowStart; rowNum <= clampedRowEnd; rowNum++) {
        const row = worksheet.getRow(rowNum);
        const rowData = [];
        
        // CRITICAL: Use index-based access for columns (NOT eachCell)
        // This ensures empty cells are preserved and column alignment matches Excel
        for (let colNum = clampedColStart; colNum <= clampedColEnd; colNum++) {
            // Get cell by column number (1-based)
            const cell = row.getCell(colNum);
            let cellValue = cell.value ?? null;
            
            // Handle different cell value types
            if (cellValue instanceof Date) {
                // Convert dates to ISO string for JSON serialization
                cellValue = cellValue.toISOString();
            } else if (typeof cellValue === 'object' && cellValue !== null) {
                // For complex objects (formulas, rich text), extract the text representation
                if (cellValue.text) {
                    cellValue = cellValue.text;
                } else if (cellValue.result !== undefined) {
                    // Formula result
                    cellValue = cellValue.result;
                } else {
                    // Fallback: stringify complex objects
                    cellValue = JSON.stringify(cellValue);
                }
            }
            
            // Always push a value (null for empty cells) to maintain column alignment
            rowData.push(cellValue);
        }
        
        windowData.push(rowData);
    }
    
    // Return windowed data with metadata
    return {
        data: windowData,
        meta: {
            totalRows: MAX_VIRTUAL_ROWS,
            totalColumns: MAX_VIRTUAL_COLS,
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

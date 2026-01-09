import { orchestratePrompt } from './promptOrchestrator.js';
import { processExcelAction, getPreviewData } from './excelService.js';
import ExcelJS from 'exceljs';

class AICommandValidationError extends Error {
  constructor(message, details) {
    super(message);
    this.name = 'AICommandValidationError';
    this.statusCode = 400;
    this.type = 'VALIDATION_ERROR';
    this.details = details;
  }
}

class AICommandParseError extends Error {
  constructor(message) {
    super(message);
    this.name = 'AICommandParseError';
    this.statusCode = 400;
    this.type = 'AI_ERROR';
  }
}

const isBlankString = (value) => typeof value === 'string' && value.trim() === '';

const requireNonEmptyString = (obj, key) => {
  const value = obj?.[key];
  if (value === undefined || value === null || isBlankString(value)) {
    throw new AICommandValidationError(`Missing required parameter "${key}".`, { key });
  }
  if (typeof value !== 'string') {
    throw new AICommandValidationError(`Parameter "${key}" must be a string.`, { key });
  }
  return value;
};

const requireNonEmptyValue = (obj, key) => {
  const value = obj?.[key];
  if (value === undefined || value === null || isBlankString(value)) {
    throw new AICommandValidationError(`Missing required parameter "${key}".`, { key });
  }
  return value;
};

const validateNoEmptyStrings = (params) => {
  if (!params || typeof params !== 'object') return;
  for (const [k, v] of Object.entries(params)) {
    if (isBlankString(v)) {
      throw new AICommandValidationError(`Parameter "${k}" must not be an empty string.`, { key: k });
    }
  }
};

const validateActionObject = (actionObj) => {
  if (!actionObj || typeof actionObj !== 'object') {
    throw new AICommandValidationError('AI output must be an object.', {});
  }

  const action = actionObj.action;
  const params = actionObj.params;

  if (!action || typeof action !== 'string' || isBlankString(action)) {
    throw new AICommandValidationError('Missing required field "action".', {});
  }

  if (!params || typeof params !== 'object') {
    throw new AICommandValidationError('Missing required field "params".', { action });
  }

  validateNoEmptyStrings(params);

  switch (action) {
    case 'ADD_COLUMN':
      requireNonEmptyString(params, 'columnName');
      requireNonEmptyString(params, 'formula');
      break;
    case 'HIGHLIGHT_ROWS':
      requireNonEmptyString(params, 'condition');
      break;
    case 'SORT_DATA': {
      requireNonEmptyString(params, 'column');
      const order = requireNonEmptyString(params, 'order').toLowerCase();
      if (order !== 'asc' && order !== 'desc') {
        throw new AICommandValidationError('Parameter "order" must be "asc" or "desc".', { key: 'order' });
      }
      params.order = order;
      break;
    }
    case 'UPDATE_ROW_VALUES': {
      requireNonEmptyString(params, 'filterColumn');
      requireNonEmptyValue(params, 'filterValue');
      const operation = requireNonEmptyString(params, 'operation');
      const allowed = new Set(['SET', '+', '-', '*', '/']);
      if (!allowed.has(operation)) {
        throw new AICommandValidationError('Parameter "operation" must be one of: SET, +, -, *, /.', { key: 'operation' });
      }
      requireNonEmptyValue(params, 'value');
      requireNonEmptyString(params, 'targetColumn');
      break;
    }
    case 'UPDATE_COLUMN_VALUES': {
      requireNonEmptyString(params, 'column');
      const operation = requireNonEmptyString(params, 'operation');
      const allowed = new Set(['SET', '+', '-', '*', '/']);
      if (!allowed.has(operation)) {
        throw new AICommandValidationError('Parameter "operation" must be one of: SET, +, -, *, /.', { key: 'operation' });
      }
      requireNonEmptyValue(params, 'value');
      break;
    }
    case 'UPDATE_KEY_VALUE':
      requireNonEmptyString(params, 'keyColumn');
      requireNonEmptyValue(params, 'keyValue');
      requireNonEmptyValue(params, 'newValue');
      if (params.valueColumn !== undefined && params.valueColumn !== null) {
        if (typeof params.valueColumn !== 'string' || isBlankString(params.valueColumn)) {
          throw new AICommandValidationError('Parameter "valueColumn" must be a non-empty string when provided.', { key: 'valueColumn' });
        }
      }
      break;
    case 'SET_CELL': {
      const cell = requireNonEmptyString(params, 'cell');
      if (!/^[A-Za-z]+[0-9]+$/.test(cell.trim())) {
        throw new AICommandValidationError('Parameter "cell" must be a valid Excel cell address (e.g. A1, B5).', { key: 'cell' });
      }
      requireNonEmptyValue(params, 'value');
      break;
    }
    case 'FIND_AND_REPLACE':
      requireNonEmptyValue(params, 'findValue');
      requireNonEmptyValue(params, 'replaceValue');
      if (params.column !== undefined && params.column !== null) {
        if (typeof params.column !== 'string' || isBlankString(params.column)) {
          throw new AICommandValidationError('Parameter "column" must be a non-empty string when provided.', { key: 'column' });
        }
      }
      break;
    case 'ERROR':
      requireNonEmptyString(actionObj, 'message');
      break;
    default:
      throw new AICommandValidationError(`Unsupported action "${action}".`, { action });
  }

  return { action, params };
};

const detectSheetContext = async (filePath, sheetId, scanDepth = 20) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = sheetId ? workbook.getWorksheet(sheetId) : workbook.worksheets[0];
  if (!worksheet) return { sheetName: sheetId || null, columns: [] };

  const maxRow = Math.min(worksheet.rowCount || scanDepth, scanDepth);
  let bestRowValues = [];
  let bestScore = 0;

  for (let r = 1; r <= maxRow; r++) {
    const row = worksheet.getRow(r);
    const texts = [];
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (v === null || v === undefined) return;
      if (typeof v === 'string') {
        const t = v.trim();
        if (t) texts.push(t);
        return;
      }
      if (typeof v === 'number') return;
      if (typeof v === 'object' && v !== null) {
        if (typeof v.text === 'string') {
          const t = v.text.trim();
          if (t) texts.push(t);
        } else if (v.result !== undefined && v.result !== null && typeof v.result === 'string') {
          const t = v.result.trim();
          if (t) texts.push(t);
        }
      }
    });

    const unique = new Set(texts.map(t => t.toLowerCase()));
    if (unique.size > bestScore) {
      bestScore = unique.size;
      bestRowValues = texts;
    }
  }

  const columns = [];
  const seen = new Set();
  for (const t of bestRowValues) {
    const key = t.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      columns.push(t);
    }
  }

  return { sheetName: worksheet.name, columns: bestScore >= 2 ? columns.slice(0, 50) : [] };
};

export const executeAICommand = async (command, filePath, sheetId) => {
  try {
    const context = await detectSheetContext(filePath, sheetId);

    const actionData = await orchestratePrompt(command, context);
    console.log('AI Action:', actionData);

    if (actionData?.action === 'ERROR') {
      throw new AICommandParseError(actionData.message || 'AI could not interpret the command.');
    }

    let actions = null;
    if (Array.isArray(actionData?.actions)) {
      actions = actionData.actions;
    } else if (actionData?.action) {
      actions = [actionData];
    } else {
      throw new AICommandParseError('AI output is missing "action" or "actions".');
    }

    if (actions.length === 0) {
      throw new AICommandParseError('AI returned an empty actions list.');
    }

    const validatedActions = actions.map(validateActionObject);

    const results = [];
    let currentFilePath = filePath;
    let anySucceeded = false;

    for (let i = 0; i < validatedActions.length; i++) {
      const step = validatedActions[i];
      try {
        const updatedFilePath = await processExcelAction(currentFilePath, step, sheetId);
        currentFilePath = updatedFilePath;
        anySucceeded = true;
        results.push({ index: i, action: step.action, success: true, filePath: updatedFilePath });
      } catch (err) {
        results.push({ index: i, action: step.action, success: false, message: err?.message || 'Execution failed' });
        break;
      }
    }

    const preview = anySucceeded ? await getPreviewData(currentFilePath) : null;

    return {
      success: results.every(r => r.success),
      action: actionData,
      results,
      filePath: anySucceeded ? currentFilePath : undefined,
      preview: preview || undefined
    };
  } catch (error) {
    console.error('AI Execution Failed:', error);
    throw error;
  }
};

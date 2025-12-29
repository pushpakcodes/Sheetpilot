import axios from 'axios';

const API_URL = 'http://localhost:5000/api';

const api = axios.create({
  baseURL: API_URL,
});

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

export const login = (email, password) => api.post('/auth/login', { email, password });
export const register = (username, email, password) => api.post('/auth/register', { username, email, password });
export const uploadFile = (formData) => api.post('/excel/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
});
export const processCommand = (command, filePath) => api.post('/excel/process', { command, filePath });
export const getFiles = () => api.get('/excel/files');

/**
 * Fetch windowed/sliced data from an Excel sheet
 * This enables virtualized rendering by fetching only visible rows/columns
 * 
 * @param {string} sheetId - File identifier (filename or path)
 * @param {number} rowStart - Starting row (1-based)
 * @param {number} rowEnd - Ending row (1-based, inclusive)
 * @param {number} colStart - Starting column (1-based)
 * @param {number} colEnd - Ending column (1-based, inclusive)
 * @param {number} sheetIndex - Sheet index (0-based, defaults to 0)
 * @returns {Promise<Object>} Windowed data with metadata
 */
/**
 * Get workbook metadata (all sheets with dimensions)
 * @param {string} workbookId - File identifier (filename or path)
 * @returns {Promise<Object>} Workbook metadata with sheets array
 */
export const getWorkbookMetadata = (workbookId) => {
    const encodedId = encodeURIComponent(workbookId);
    return api.get(`/excel/workbook/${encodedId}/metadata`);
};

/**
 * Fetch windowed/sliced data from an Excel sheet
 * This enables virtualized rendering by fetching only visible rows/columns
 * 
 * @param {string} workbookId - File identifier (filename or path)
 * @param {number} rowStart - Starting row (1-based)
 * @param {number} rowEnd - Ending row (1-based, inclusive)
 * @param {number} colStart - Starting column (1-based)
 * @param {number} colEnd - Ending column (1-based, inclusive)
 * @param {string} sheetId - Sheet name (NOT index)
 * @returns {Promise<Object>} Windowed data with metadata
 */
export const getSheetWindow = (workbookId, rowStart, rowEnd, colStart, colEnd, sheetId) => {
    const encodedId = encodeURIComponent(workbookId);
    return api.get(`/excel/workbook/${encodedId}/window`, {
        params: {
            rowStart,
            rowEnd,
            colStart,
            colEnd,
            sheetId // Use sheet name, not index
        }
    });
};

/**
 * Update a single cell in the workbook
 * @param {string} workbookId - File identifier (filename or path)
 * @param {string} sheetId - Sheet name (NOT index)
 * @param {number} row - Row number (1-based)
 * @param {number} col - Column number (1-based)
 * @param {any} value - New cell value
 * @returns {Promise<Object>} Success response
 */
export const updateCell = (workbookId, sheetId, row, col, value) => {
    const encodedId = encodeURIComponent(workbookId);
    return api.post(`/excel/workbook/${encodedId}/cell`, {
        sheetId, // Use sheet name, not index
        row,
        col,
        value
    });
};

export default api;

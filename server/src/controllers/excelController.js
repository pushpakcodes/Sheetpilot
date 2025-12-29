import User from '../models/User.js';
import { executeAICommand } from '../services/aiService.js';
import { getPreviewData, getWindowedSheetData, getWorkbookMetadata, updateCell } from '../services/excelService.js';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export const uploadExcel = async (req, res) => {
  if (!req.file) {
      return res.status(400).send({ message: 'Please upload a file' });
  }
  
  // If user is logged in, save file ref to user
  if (req.user) {
      const user = await User.findById(req.user._id);
      user.files.push({
          originalName: req.file.originalname,
          filename: req.file.filename,
          path: req.file.path
      });
      await user.save();
  }

  const preview = await getPreviewData(req.file.path);

  // Normalize path separators to forward slashes for cross-platform compatibility
  // This ensures the path works the same on Windows, Linux, and Mac
  const normalizedPath = req.file.path.replace(/\\/g, '/');

  res.send({
      message: 'File uploaded successfully',
      filePath: normalizedPath,
      filename: req.file.filename,
      preview
  });
};

export const processCommand = async (req, res) => {
    const { command, filePath } = req.body;
    
    if (!command || !filePath) {
        return res.status(400).json({ message: 'Command and filePath are required' });
    }

    try {
        const result = await executeAICommand(command, filePath);
        res.json(result);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
};

export const getUserFiles = async (req, res) => {
    if (!req.user) {
        return res.status(401).json({ message: 'Not authorized' });
    }
    const user = await User.findById(req.user._id);
    res.json(user.files);
};

/**
 * Get windowed/sliced data from an Excel sheet
 * This endpoint enables virtualized rendering by returning only a slice of rows/columns
 * 
 * Query params:
 * - rowStart: Starting row (1-based, Excel convention)
 * - rowEnd: Ending row (1-based, inclusive)
 * - colStart: Starting column (1-based, Excel convention)
 * - colEnd: Ending column (1-based, inclusive)
 * - sheetIndex: Sheet index (0-based, defaults to 0)
 * 
 * The sheetId can be:
 * - A filename (if file is in uploads folder)
 * - A full file path (for uploaded files)
 */
/**
 * Get workbook metadata (all sheets with dimensions)
 * GET /api/excel/workbook/:workbookId/metadata
 */
export const getWorkbookMetadataController = async (req, res) => {
    try {
        let { workbookId } = req.params;
        workbookId = decodeURIComponent(workbookId);
        
        // Resolve file path
        let filePath = resolveFilePath(workbookId);
        
        if (!filePath) {
            return res.status(404).json({ 
                message: `Workbook not found: ${workbookId}` 
            });
        }
        
        const metadata = await getWorkbookMetadata(filePath);
        res.json(metadata);
        
    } catch (error) {
        console.error('Error fetching workbook metadata:', error);
        res.status(500).json({ 
            message: 'Error fetching workbook metadata', 
            error: error.message 
        });
    }
};

/**
 * Update a single cell in the workbook
 * POST /api/excel/workbook/:workbookId/cell
 * Body: { sheetIndex, row, col, value }
 */
export const updateCellController = async (req, res) => {
    try {
        let { workbookId } = req.params;
        workbookId = decodeURIComponent(workbookId);
        const { sheetId, row, col, value } = req.body;
        
        // Validate required parameters
        if (!sheetId || !row || !col || value === undefined) {
            return res.status(400).json({ 
                message: 'Missing required parameters: sheetId (sheet name), row, col, value' 
            });
        }
        
        // Validate numeric parameters
        const parsedRow = parseInt(row, 10);
        const parsedCol = parseInt(col, 10);
        
        if (isNaN(parsedRow) || isNaN(parsedCol)) {
            return res.status(400).json({ 
                message: 'Invalid parameters: row and col must be numbers' 
            });
        }
        
        if (parsedRow < 1 || parsedCol < 1) {
            return res.status(400).json({ 
                message: 'Row and column must be >= 1 (Excel uses 1-based indexing)' 
            });
        }
        
        // Resolve file path
        let filePath = resolveFilePath(workbookId);
        
        if (!filePath) {
            return res.status(404).json({ 
                message: `Workbook not found: ${workbookId}` 
            });
        }
        
        // Update the cell using sheet name (NOT index)
        await updateCell(filePath, sheetId, parsedRow, parsedCol, value);
        
        res.json({ 
            success: true, 
            message: 'Cell updated successfully' 
        });
        
    } catch (error) {
        console.error('Error updating cell:', error);
        res.status(500).json({ 
            message: 'Error updating cell', 
            error: error.message 
        });
    }
};

/**
 * Helper function to resolve file path from workbookId
 * Handles various path formats and normalizes them
 */
const resolveFilePath = (workbookId) => {
    // Normalize path separators - convert all to forward slashes first, then to OS-specific
    const normalizedId = workbookId.replace(/\\/g, '/').replace(/\//g, path.sep);
    
    const uploadsDirRelative = path.join(process.cwd(), 'uploads');
    const uploadsDirAbsolute = path.join(__dirname, '../../uploads');
    
    // Strategy 1: Check if it's already an absolute path that exists
    if (path.isAbsolute(normalizedId) && fs.existsSync(normalizedId)) {
        return normalizedId;
    }
    
    // Strategy 2: Check if it's a relative path from current working directory
    if (fs.existsSync(normalizedId)) {
        return path.resolve(normalizedId);
    }
    
    // Strategy 3: If it contains 'uploads', try reconstructing the path
    if (normalizedId.includes('uploads')) {
        // Extract the part after 'uploads/'
        const parts = normalizedId.split('uploads');
        const afterUploads = parts[parts.length - 1].replace(/^[\/\\]/, ''); // Remove leading slash
        
        // Try with relative uploads dir
        const relativePath = path.join(uploadsDirRelative, afterUploads);
        if (fs.existsSync(relativePath)) {
            return relativePath;
        }
        
        // Try with absolute uploads dir
        const absolutePath = path.join(uploadsDirAbsolute, afterUploads);
        if (fs.existsSync(absolutePath)) {
            return absolutePath;
        }
    }
    
    // Strategy 4: Try as filename in uploads directory
    const filename = path.basename(normalizedId);
    const filenameRelative = path.join(uploadsDirRelative, filename);
    if (fs.existsSync(filenameRelative)) {
        return filenameRelative;
    }
    
    const filenameAbsolute = path.join(uploadsDirAbsolute, filename);
    if (fs.existsSync(filenameAbsolute)) {
        return filenameAbsolute;
    }
    
    // Strategy 5: Try with relative path from cwd
    const relativePath = path.join(uploadsDirRelative, normalizedId);
    if (fs.existsSync(relativePath)) {
        return relativePath;
    }
    
    // Strategy 6: Try with absolute path from server directory
    const absolutePath = path.join(uploadsDirAbsolute, normalizedId);
    if (fs.existsSync(absolutePath)) {
        return absolutePath;
    }
    
    // Log all attempted paths for debugging
    console.error('❌ File resolution failed for:', workbookId);
    console.error('Attempted paths:', {
        normalized: normalizedId,
        relativePath,
        absolutePath,
        filenameRelative,
        filenameAbsolute,
        cwd: process.cwd(),
        __dirname: __dirname
    });
    
    return null;
};

export const getSheetWindow = async (req, res) => {
    try {
        // Handle both workbookId (new) and sheetId (legacy) parameter names
        let workbookId = req.params.workbookId || req.params.sheetId;
        workbookId = decodeURIComponent(workbookId);
        
        // CRITICAL: Use sheetId (name) instead of sheetIndex
        const { rowStart, rowEnd, colStart, colEnd, sheetId } = req.query;
        
        if (!sheetId) {
            return res.status(400).json({ 
                message: 'Missing required query parameter: sheetId (sheet name)' 
            });
        }
        
        // Validate required query parameters
        if (!rowStart || !rowEnd || !colStart || !colEnd) {
            return res.status(400).json({ 
                message: 'Missing required query parameters: rowStart, rowEnd, colStart, colEnd' 
            });
        }
        
        // Parse and validate numeric parameters
        const parsedRowStart = parseInt(rowStart, 10);
        const parsedRowEnd = parseInt(rowEnd, 10);
        const parsedColStart = parseInt(colStart, 10);
        const parsedColEnd = parseInt(colEnd, 10);
        
        if (isNaN(parsedRowStart) || isNaN(parsedRowEnd) || 
            isNaN(parsedColStart) || isNaN(parsedColEnd)) {
            return res.status(400).json({ 
                message: 'Invalid query parameters: rowStart, rowEnd, colStart, colEnd must be numbers' 
            });
        }
        
        // Validate range
        if (parsedRowStart > parsedRowEnd || parsedColStart > parsedColEnd) {
            return res.status(400).json({ 
                message: 'Invalid range: start must be <= end' 
            });
        }
        
        if (parsedRowStart < 1 || parsedColStart < 1) {
            return res.status(400).json({ 
                message: 'Invalid range: start values must be >= 1 (Excel uses 1-based indexing)' 
            });
        }
        
        // Resolve file path using helper function
        const filePath = resolveFilePath(workbookId);
        
        if (!filePath) {
            console.error(`File not found: ${workbookId}`);
            console.error('Attempted paths:', {
                original: workbookId,
                uploadsRelative: path.join(process.cwd(), 'uploads', workbookId),
                uploadsAbsolute: path.join(__dirname, '../../uploads', workbookId)
            });
            return res.status(404).json({ 
                message: `File not found: ${workbookId}` 
            });
        }
        
        // Fetch windowed data using sheet name (NOT index)
        const result = await getWindowedSheetData(
            filePath,
            sheetId, // Use sheet name
            parsedRowStart,
            parsedRowEnd,
            parsedColStart,
            parsedColEnd
        );
        
        // Return the windowed data with metadata
        // The frontend uses this to render only visible cells
        res.json(result);
        
    } catch (error) {
        console.error('Error fetching sheet window:', error);
        res.status(500).json({ 
            message: 'Error fetching sheet window', 
            error: error.message 
        });
    }
};

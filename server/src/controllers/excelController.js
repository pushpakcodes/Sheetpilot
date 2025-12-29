import User from '../models/User.js';
import { executeAICommand } from '../services/aiService.js';
import { getPreviewData, getWindowedSheetData } from '../services/excelService.js';
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

  res.send({
      message: 'File uploaded successfully',
      filePath: req.file.path,
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
export const getSheetWindow = async (req, res) => {
    try {
        // Decode URL-encoded sheetId (Express usually does this automatically, but be explicit)
        let { sheetId } = req.params;
        sheetId = decodeURIComponent(sheetId);
        const { rowStart, rowEnd, colStart, colEnd, sheetIndex = 0 } = req.query;
        
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
        const parsedSheetIndex = parseInt(sheetIndex, 10);
        
        if (isNaN(parsedRowStart) || isNaN(parsedRowEnd) || 
            isNaN(parsedColStart) || isNaN(parsedColEnd) || 
            isNaN(parsedSheetIndex)) {
            return res.status(400).json({ 
                message: 'Invalid query parameters: all values must be numbers' 
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
        
        // Resolve file path
        // sheetId can be a filename or a full path
        let filePath;
        
        // Get uploads directory - multer saves to 'uploads/' relative to process.cwd()
        // We need to check both possible locations:
        // 1. Relative to process.cwd() (if server runs from server/ directory)
        // 2. Relative to server directory (if server runs from project root)
        const uploadsDirRelative = path.join(process.cwd(), 'uploads');
        const uploadsDirAbsolute = path.join(__dirname, '../../uploads');
        
        // Check if it's a full path first
        if (fs.existsSync(sheetId)) {
            filePath = sheetId;
        } else {
            // Try relative to process.cwd() first (multer's default behavior)
            const relativePath = path.join(uploadsDirRelative, sheetId);
            if (fs.existsSync(relativePath)) {
                filePath = relativePath;
            } else {
                // Try absolute path from server directory
                const absolutePath = path.join(uploadsDirAbsolute, sheetId);
                if (fs.existsSync(absolutePath)) {
                    filePath = absolutePath;
                } else {
                    console.error(`File not found: ${sheetId}`);
                    console.error(`Checked paths: ${relativePath}, ${absolutePath}`);
                    return res.status(404).json({ 
                        message: `File not found: ${sheetId}` 
                    });
                }
            }
        }
        
        // Fetch windowed data
        // This loads the full Excel file but returns only the requested slice
        const result = await getWindowedSheetData(
            filePath,
            parsedSheetIndex,
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

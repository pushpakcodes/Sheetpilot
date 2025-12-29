import express from 'express';
import { 
    uploadExcel, 
    processCommand, 
    getUserFiles, 
    getSheetWindow,
    getWorkbookMetadataController,
    updateCellController
} from '../controllers/excelController.js';
import upload from '../middleware/uploadMiddleware.js';
import { protect, extractUser } from '../middleware/authMiddleware.js';

const router = express.Router();

router.post('/upload', extractUser, upload.single('file'), uploadExcel);
router.post('/process', extractUser, processCommand);
router.get('/files', protect, getUserFiles);

// Workbook metadata endpoint
// GET /api/excel/workbook/:workbookId/metadata
router.get('/workbook/:workbookId/metadata', extractUser, getWorkbookMetadataController);

// Windowed data endpoint for virtualized rendering
// GET /api/excel/workbook/:workbookId/window?rowStart=1&rowEnd=100&colStart=1&colEnd=30&sheetIndex=0
router.get('/workbook/:workbookId/window', extractUser, getSheetWindow);

// Legacy endpoint for backward compatibility
// GET /api/excel/sheets/:sheetId/window?rowStart=1&rowEnd=100&colStart=1&colEnd=30&sheetIndex=0
router.get('/sheets/:sheetId/window', extractUser, getSheetWindow);

// Cell update endpoint
// POST /api/excel/workbook/:workbookId/cell
router.post('/workbook/:workbookId/cell', extractUser, updateCellController);

export default router;

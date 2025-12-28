import express from 'express';
import { uploadExcel, processCommand, getUserFiles, getSheetWindow } from '../controllers/excelController.js';
import upload from '../middleware/uploadMiddleware.js';
import { protect, extractUser } from '../middleware/authMiddleware.js';

const router = express.Router();

router.post('/upload', extractUser, upload.single('file'), uploadExcel);
router.post('/process', extractUser, processCommand);
router.get('/files', protect, getUserFiles);
// Windowed data endpoint for virtualized rendering
// GET /api/excel/sheets/:sheetId/window?rowStart=1&rowEnd=100&colStart=1&colEnd=30&sheetIndex=0
router.get('/sheets/:sheetId/window', extractUser, getSheetWindow);

export default router;

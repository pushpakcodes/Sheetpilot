import express from 'express';
import { uploadExcel, processCommand, getUserFiles } from '../controllers/excelController.js';
import upload from '../middleware/uploadMiddleware.js';
import { protect, extractUser } from '../middleware/authMiddleware.js';

const router = express.Router();

router.post('/upload', extractUser, upload.single('file'), uploadExcel);
router.post('/process', extractUser, processCommand);
router.get('/files', protect, getUserFiles);

export default router;

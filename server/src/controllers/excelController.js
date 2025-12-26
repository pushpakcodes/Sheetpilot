import User from '../models/User.js';
import { executeAICommand } from '../services/aiService.js';
import { getPreviewData } from '../services/excelService.js';
import path from 'path';

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

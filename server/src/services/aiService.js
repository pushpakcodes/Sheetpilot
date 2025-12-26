import { orchestratePrompt } from './promptOrchestrator.js';
import { processExcelAction, getPreviewData } from './excelService.js';

export const executeAICommand = async (command, filePath) => {
  try {
    // 1. Convert command to JSON
    const actionData = await orchestratePrompt(command);
    console.log('AI Action:', actionData);

    if (actionData.action === 'ERROR') {
      throw new Error(actionData.message);
    }

    // 2. Execute Action on Excel File
    const updatedFilePath = await processExcelAction(filePath, actionData);

    // 3. Get Preview Data
    const preview = await getPreviewData(updatedFilePath);

    return {
      success: true,
      action: actionData,
      filePath: updatedFilePath,
      preview
    };
  } catch (error) {
    console.error('AI Execution Failed:', error);
    throw error;
  }
};

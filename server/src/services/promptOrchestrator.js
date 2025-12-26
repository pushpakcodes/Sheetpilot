import OpenAI from 'openai';
import dotenv from 'dotenv';

dotenv.config();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const SYSTEM_PROMPT = `
You are an AI assistant for an Excel processing application called SheetPilot.
Your goal is to interpret natural language commands and convert them into structured JSON actions that can be executed by the system.
The system uses ExcelJS to manipulate the workbook.

Supported Actions:
- ADD_COLUMN: Add a new column with a formula or value.
- HIGHLIGHT_ROWS: Highlight rows based on a condition.
- CREATE_PIVOT_TABLE: Create a pivot table.
- SORT_DATA: Sort data by a specific column.
- ADD_ROWS: Add empty rows.
- SUMMARY_STATS: Generate summary statistics.
- CONDITIONAL_FORMAT: Apply conditional formatting.

Output Format:
You must return a JSON object (and ONLY a JSON object) with the following structure:
{
  "action": "ACTION_NAME",
  "params": {
    // Action specific parameters
  }
}

Examples:
1. Input: "Add a new column called Profit = Revenue - Cost"
   Output:
   {
     "action": "ADD_COLUMN",
     "params": {
       "columnName": "Profit",
       "formula": "Revenue - Cost"
     }
   }

2. Input: "Highlight rows where sales > 50,000"
   Output:
   {
     "action": "HIGHLIGHT_ROWS",
     "params": {
       "condition": "sales > 50000",
       "color": "FFFF00"
     }
   }

3. Input: "Sort data by date descending"
   Output:
   {
     "action": "SORT_DATA",
     "params": {
       "column": "date",
       "order": "desc"
     }
   }

If the command is unclear or not supported, return:
{
  "action": "ERROR",
  "message": "Reason why the command is invalid"
}
`;

export const orchestratePrompt = async (userCommand) => {
  try {
    const response = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userCommand },
      ],
      temperature: 0,
    });

    const content = response.choices[0].message.content;
    try {
        return JSON.parse(content);
    } catch (e) {
        // Fallback if LLM wraps in markdown code block
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            return JSON.parse(jsonMatch[0]);
        }
        throw new Error("Failed to parse LLM response");
    }
  } catch (error) {
    console.error("AI Orchestration Error:", error);
    throw error;
  }
};

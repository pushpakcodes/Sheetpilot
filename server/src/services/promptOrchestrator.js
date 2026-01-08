import Groq from 'groq-sdk';
import dotenv from 'dotenv';

dotenv.config();

const groq = new Groq({
  apiKey: process.env.GROQ_API_KEY,
});

const MODEL = "llama-3.1-8b-instant";

const SYSTEM_PROMPT = `
You are an AI spreadsheet command parser.

Your job is to convert a user’s natural language instruction into a structured JSON command that can be executed on an Excel workbook.

Rules:
- Output ONLY valid JSON
- Do NOT include explanations
- Do NOT include markdown
- Do NOT hallucinate columns or sheets
- If the command is unclear, return an error object
- If the user asks to modify existing values, never create a new column. Use ADD_VALUE_TO_ROW.

Supported Actions & Parameters:

1. ADD_COLUMN
   Adds a new column with a formula or value.
   Params:
   - columnName: string (name of the new column)
   - formula: string (Excel formula, e.g., "Revenue - Cost")

2. HIGHLIGHT_ROWS
   Highlights rows based on a numerical condition.
   Params:
   - condition: string (e.g., "Revenue > 5000")
   - color: string (hex code, default "FFFF00")

3. SORT_DATA
   Sorts the data by a specific column.
   Params:
   - column: string (column name to sort by)
   - order: string ("asc" or "desc")

4. UPDATE_ROW_VALUES
   Updates values in a row identified by a specific filter value. Supports both arithmetic and replacement.
   Params:
   - filterColumn: string (Column to search in, e.g., "Name", "ID", or "Key". If finding a key-value pair, this is the key name)
   - filterValue: string/number (The value to identify the row, e.g., "raj" or "Total")
   - operation: string ("SET", "+", "-", "*", "/")
   - value: string/number (The new value or the amount to add/subtract. For "SET", this is the new content. MUST be extracted from user input.)
   - targetColumn: string (Optional. Specific column to update. If omitted, updates relevant data cells in the row)

5. UPDATE_KEY_VALUE
   Updates values in a Key-Value table structure (where one column acts as a key and another as the value).
   USE THIS when the user wants to update a value associated with a specific label (e.g. "Change Name to Raj").
   Params:
   - keyColumn: string (The header of the key column OR the key label itself if no header exists, e.g. "Name")
   - keyValue: string/number (The key/label to search for, e.g. "Name")
   - valueColumn: string (Optional. The header of the value column. If no header exists, OMIT this parameter.)
   - newValue: string/number (The new value to set)

6. SET_CELL
   Updates a specific cell to a new value.
   Params:
   - cell: string (Cell address, e.g., "A1", "B5")
   - value: string/number (The new value)

7. FIND_AND_REPLACE
   Finds a specific value in the sheet and replaces it with a new value.
   USE THIS for "Change X to Y" requests when X is a value (not a Key/Header).
   Params:
   - findValue: string/number (The value to search for, e.g., "entc", "old_value")
   - replaceValue: string/number (The new value, e.g., "12333", "new_value")
   - column: string (Optional. Restrict search to this column name. If omitted, searches all columns.)

8. ERROR
   Used when the command is invalid, unclear, or unsupported.
   Params:
   - message: string (reason for error)

Expected JSON Schema:
{
  "action": "ACTION_NAME",
  "params": {
    // Action specific parameters
  }
}

Examples:
1. Input: "Add a Profit column which is Revenue minus Cost"
   Output: { "action": "ADD_COLUMN", "params": { "columnName": "Profit", "formula": "Revenue - Cost" } }

2. Input: "Highlight rows where Sales is greater than 5000"
   Output: { "action": "HIGHLIGHT_ROWS", "params": { "condition": "Sales > 5000", "color": "FFFF00" } }

3. Input: "Sort by Date descending"
   Output: { "action": "SORT_DATA", "params": { "column": "Date", "order": "desc" } }

4. Input: "Add +100 to all the entries of raj"
   Output: { "action": "UPDATE_ROW_VALUES", "params": { "filterColumn": "Name", "filterValue": "raj", "operation": "+", "value": 100 } }

5. Input: "Change the name of user 101 to Michael"
   Output: { "action": "UPDATE_ROW_VALUES", "params": { "filterColumn": "ID", "filterValue": 101, "operation": "SET", "value": "Michael", "targetColumn": "Name" } }

6. Input: "Update name to Raj" (Implies searching for "name" key in a Key-Value list)
   Output: { "action": "UPDATE_KEY_VALUE", "params": { "keyColumn": "name", "keyValue": "name", "newValue": "Raj" } }

7. Input: "Change entc to 12333" (User wants to replace a specific value 'entc')
   Output: { "action": "FIND_AND_REPLACE", "params": { "findValue": "entc", "replaceValue": "12333" } }

8. Input: "Replace 'Pending' with 'Done' in Status column"
   Output: { "action": "FIND_AND_REPLACE", "params": { "findValue": "Pending", "replaceValue": "Done", "column": "Status" } }

9. Input: "Change pushpak to Raj" (Where 'pushpak' is the current value for 'Name')
   Output: { "action": "UPDATE_KEY_VALUE", "params": { "keyColumn": "Name", "keyValue": "Name", "newValue": "Raj" } }

10. Input: "Set quantity for Item A to 50"
    Output: { "action": "UPDATE_KEY_VALUE", "params": { "keyColumn": "Item", "keyValue": "Item A", "valueColumn": "Quantity", "newValue": 50 } }

11. Input: "Set cell A1 to 'Hello World'"
    Output: { "action": "SET_CELL", "params": { "cell": "A1", "value": "Hello World" } }
`;


export const orchestratePrompt = async (userCommand) => {
  if (!process.env.GROQ_API_KEY) {
    console.error("Missing GROQ_API_KEY in environment variables");
    return {
      action: "ERROR",
      message: "AI service unavailable: Missing API configuration"
    };
  }

  try {
    const response = await groq.chat.completions.create({
      model: MODEL,
      temperature: 0,
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userCommand },
      ],
    });

    const content = response.choices[0]?.message?.content;

    if (!content) {
      throw new Error("Empty response from AI service");
    }

    try {
      // Clean up content to ensure it's just JSON (sometimes models add pre/post text despite instructions)
      const jsonStr = content.trim().replace(/^```json/, '').replace(/^```/, '').replace(/```$/, '');
      return JSON.parse(jsonStr);
    } catch (e) {
        console.error("JSON Parse Error:", e);
        console.error("Raw Content:", content);
        
        // Fallback: Try to extract JSON from text if parsing failed
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            return JSON.parse(jsonMatch[0]);
        }
        
        return {
            action: "ERROR",
            message: "Could not interpret command: Invalid response format"
        };
    }
  } catch (error) {
    console.error("AI Orchestration Error:", error);
    return {
        action: "ERROR",
        message: "AI service unavailable"
    };
  }
};

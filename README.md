# SheetPilot - Cursor for Excel

SheetPilot is a full-stack AI-powered web application that allows users to perform Excel-like operations using natural language commands.

## Features

- **Natural Language Processing**: Type commands like "Add a Profit column" or "Highlight rows where Sales > 50000".
- **Real-time Preview**: See changes instantly in the web interface.
- **ExcelJS Integration**: Robust Excel file manipulation on the backend.
- **Secure Authentication**: JWT-based login and signup.
- **Modern UI**: Built with React, Tailwind CSS, and Framer Motion.

## Tech Stack

- **Frontend**: React, Vite, Tailwind CSS, Framer Motion, Axios
- **Backend**: Node.js, Express, MongoDB, ExcelJS, OpenAI API
- **AI**: OpenAI GPT-3.5/4

## Setup Instructions

### Prerequisites

- Node.js (v18+)
- MongoDB (running locally or a connection string)
- OpenAI API Key

### Installation

1.  **Clone the repository** (if applicable)

2.  **Install Dependencies**
    Run the following command in the root directory to install dependencies for both client and server:
    ```bash
    npm install
    npm run install-all
    ```
    (Or manually `cd client && npm install` and `cd server && npm install`)

3.  **Environment Configuration**

    Create a `.env` file in the `server` directory:
    ```env
    PORT=5000
    MONGO_URI=mongodb://localhost:27017/sheetpilot
    JWT_SECRET=your_jwt_secret
    OPENAI_API_KEY=your_openai_api_key
    ```

4.  **Run the Application**

    From the root directory:
    ```bash
    npm start
    ```
    This will start both the backend server (port 5000) and the frontend client (port 5173).

5.  **Access the App**
    Open [http://localhost:5173](http://localhost:5173) in your browser.

## Usage

1.  Register/Login to your account.
2.  Upload an Excel file (.xlsx).
3.  Type a command in the input box (e.g., "Sort by Revenue descending").
4.  Wait for AI to process and view the updated preview.
5.  Download the modified file.

## AI Commands Supported

- Add new columns with formulas
- Highlight rows based on conditions
- Sort data
- (More to come)

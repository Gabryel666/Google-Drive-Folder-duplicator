# Google Drive Duplicator

This project allows you to duplicate a complete folder structure from a Google Drive folder (shared or not) to your own Drive.

## Features

*   **Recursive Copy**: Copies all files and subfolders.
*   **Time Limit Management**: Handles the 6-minute execution time limit of Google Apps Script. If the copy isn't finished, you can resume it, and it will pick up exactly where it left off.
*   **Verification**: Compares the file count between the source and destination to ensure integrity.
*   **Google Sheets Interface**: Easy control via a spreadsheet.

## Installation

### Option A: Copy-Paste (Simple)

1.  Create a new **Google Sheet**.
2.  Go to **Extensions** > **Apps Script**.
3.  Delete existing code in `Code.gs`.
4.  Copy the content of the `src/Code.js` file from this repository and paste it into the editor.
5.  Save.
6.  Reload your Google Sheet. A "Drive Duplicator" menu will appear.

### Option B: Using CLASP (Advanced)

If you have Node.js installed:

1.  Install clasp: `npm install -g @google/clasp`
2.  Login: `clasp login`
3.  Create a sheet: `clasp create --type sheets --title "Drive Duplicator"` (or clone an existing project).
4.  Push the code: `clasp push`

## Google Sheet Configuration

The script expects the following columns (order isn't strict, but headers help):

| Row 1 | A | B | C | D |
| :--- | :--- | :--- | :--- | :--- |
| **Headers** | **Source Folder ID** | **Status** | **Destination URL** | **Verification** |

*   **Source Folder ID**: The ID of the folder you want to copy (the part at the end of the folder URL).
*   **Status**: Leave empty initially. The script will update it to "Pending", "Processing", "Done", or "Error".
*   **Destination URL**: Populated by the script.
*   **Verification**: Populated by the verification tool.

## Usage

1.  Fill in the Source Folder ID in column A.
2.  Go to the menu **Drive Duplicator** > **Start Copy**.
3.  If the script stops (time limit), the status will remain "Processing" (or indicate "Time Limit"). Simply click **Start Copy** again to resume.
4.  Once finished, the status will be "Done".
5.  To verify, click **Drive Duplicator** > **Verify Folder**.

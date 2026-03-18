# Google Sheets Jira-Style Board

This project is a Google Apps Script web app that turns a spreadsheet into a draggable 3-lane task board.

## Spreadsheet

The script is configured for this spreadsheet:

- `1K0c-r44qLJz4bOPrbRsCW1sOvHOYZ9SOvaBNQKkSa9o`

Required tabs:

- `List of Task`
- `Completed Tasks/For Monitoring`
- `Cancelled`

Required headers in exact order:

- `ID`
- `COMPANY`
- `DESCRIPTION`
- `STATUS`
- `REMARKS`
- `NOTES`
- `Links`
- `Progress`

If the tabs are empty, running `initializeBoard()` or opening the app will create the headers automatically. If your current tabs already use the original 7 columns without `ID`, the script will insert a hidden `ID` column at the front and keep your existing task data intact. Tabs with any other header layout will throw a validation error to protect your data.

## Files

- `Code.gs`: server logic, authorization, sheet validation, row move/update functions
- `Index.html`: main HTML shell
- `Styles.html`: board and modal styles
- `Scripts.html`: drag/drop and modal client logic
- `appsscript.json`: Apps Script manifest

## Deploy

1. Create a new Apps Script project or bind one to the target spreadsheet.
2. Copy these files into the Apps Script project.
3. Deploy as a web app.
4. Recommended deployment:
   - Execute as: `Me`
   - Who has access: your Google Workspace/domain or allowed signed-in users
5. After deployment, set script property `WEB_APP_URL` to the deployed URL if the custom menu should open the latest deployment reliably.

## Notes

- The script uses `LockService` for move and update operations to reduce duplicate writes during simultaneous usage.
- Authorization checks currently allow the spreadsheet owner and spreadsheet editors.
- Moving a card updates its `STATUS` automatically based on the destination lane:
  - `List of Task` -> `TODO`
  - `Completed Tasks/For Monitoring` -> `DONE`
  - `Cancelled` -> `CANCELLED`

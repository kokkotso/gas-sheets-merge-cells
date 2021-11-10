# Merge Cells Add-On for Google Sheets
This is a Google Apps Script add-on for Google Sheets to automatically merge cells based on user-provided criteria.

## Installation & Configuration
This script is not yet ready for publication (and therefore, installation via the Google Sheets add-on marketplace). To use it, follow the below steps:

1. Create a new [Google Sheets spreadsheet](https://docs.google.com/spreadsheets), or open an existing one.
2. In the menu bar, go to **Extensions** > **Apps Script**
3. Copy-paste the code from `index.js` into the editor (make sure to delete the default `myFunction` declaration in the editor first) and save.
4. In the sidebar menu, go to **Triggers** and create a new trigger that runs the function `onOpen`on the event type `On Open`.
5. Go back to the spreadsheet and refresh the browser window. After the sheet full loads, you should see an additional item in the nav bar named "Merge Cells". Clicking on the "Run Merge" option in the "Merge Cells" submenu starts the script. 

**Note:** When saving the trigger, you may encounter an error if you have pop-ups blocked in your browser settings. Enable pop-ups from Google Apps Script in order to see the authorization dialog. You must authorize the script to run on your Google Account before you can use it in the spreadsheet.

## Usage
Currently, the script only works for two column spreadsheets and merges cells in column A based on similarities in column B. The script is non-destructive and does not mutate the source sheet. Instead, it generates a new sheet within the spreadsheet with the cells merged.

### Use Case
This script was built to help automate the process of transforming spreadsheets mapping 301 redirects into the format needed for upload into Akamai as part of a global websites migration at my employer, Stanley Black & Decker, Inc.

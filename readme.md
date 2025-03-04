# iRacing Calendar Spreadsheet Creator

## Requirements

- [Powershell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5)
-- Tested on Powershell 7.5
- Excel or another spreadsheet app
- iRacing Series JSON file (will try to provide in repo, but can be acquired through iRacing's API)
-- I am unsure if I am legally allowed to redistribute their data... will ask for forgiveness later if so

How to use the program:
1. Run the iRacing Calendar program via the terminal or executable
2. Choose JSON file if needed.
3. Click items in the left list box to see the tracks in that series.
4. Click the checkbox to add that series to your calendar.
5. Repeat as needed. 8 series are recommened for the small template.
6. If you'd prefer to output all the series, click that checkbox.
7. Choose output filename and path.
8. Click Create CSV

How to create the spreadsheet:
1. Open the created CSV in Excel.
2. As of March 3 2025, there is a formatting bug and the output places all the series into column B, comma-delimited. Step 3 explains how to fix this.
3. Use the "Text to Columns" feature in Excel to format that column properly. See "Text to Columns" appendix if needed.
-- Data > Data Tools section > Text to Columns
4. Open the template file you want to use.
5. As of March 3 2025, there are two tabs or sheets for the small and large template.
6. Once properly formatted, copy the columns you want in the template. 8 series recommended for small template.
7. Click back into the Excel template and click cell B1.
8. Use the "Paste Special" feature and choose to copy "Values". The keyboard shortcut Ctrl+Shift+V may work.
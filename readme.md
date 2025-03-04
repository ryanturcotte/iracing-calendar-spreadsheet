# iRacing Calendar Spreadsheet Creator

## Requirements

- [Powershell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5)
-- Tested on Powershell 7.5
- Excel or another spreadsheet app
- iRacing Series JSON file (will try to provide in repo, but can be acquired through iRacing's API)
-- I am unsure if I am legally allowed to redistribute their data... will ask for forgiveness later if so

## How to use the program:
1. Run the iRacing Calendar program via the terminal or executable
2. Choose JSON file if needed.
3. Click items in the left list box to see the tracks in that series.
4. Click the checkbox to add that series to your calendar.
5. Repeat as needed. 8 series are recommened for the small template.
6. If you'd prefer to output all the series, click that checkbox.
7. Choose output filename and path.
8. Click Create CSV

## How to create the spreadsheet:
1. Open the created CSV in Excel.
2. Column B will be all of the track series information. There is a formatting bug and it does not properly format the CSV. Select Column B from row 1 to row 16
3. Use the "Text to Columns" feature in Excel to format that column properly. See "Text to Columns" appendix if needed.
-- Data > Data Tools section > Text to Columns
4. Open the template file you want to use. This is TemplateAll.xlsx
5. As of March 3 2025, there are two tabs or sheets for the small and large template.
6. Once properly formatted, copy the columns you want in the template. 8 series recommended for small template.
7. Click back into the Excel template and click cell B1.
8. Use the "Paste Special" feature and choose to copy "Values". The keyboard shortcut Ctrl+Shift+V may work.

##Appendix

### How to use "Text to Columns"
https://support.microsoft.com/en-us/office/split-text-into-different-columns-with-the-convert-text-to-columns-wizard-30b14928-5550-41f5-97ca-7a3e9c363ed7

TODO

##Known Issues

1. iRacing's "time" values don't all use the same formatting, so some series may not be minimized to fit the time cell
2. Series that run more than 12 weeks or less than 12 weeks are skipped. At the moment I have no interest in keeping track of them myself, and they complicate the table if accounting for them.
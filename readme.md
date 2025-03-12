# iRacing Calendar Spreadsheet Creator

This PowerShell program scrapes a JSON of iRacing season series data and presents a GUI of all the series available, with the goal to output selected ones to a CSV. A XLSX template is provided that can present 8 series on 1 landscape page. By copying data from the CSV to the included XLSX template, you can build a 12 week iRacing schedule fit for printing in a matter of minutes.

## Requirements

- [Powershell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.5)
    - Tested on Powershell 7.5
- Excel or another spreadsheet app
    - Tested with Excel 365 Version 2501
- iRacing Series JSON file (provided in jsons folder, but can be acquired through iRacing's API)
    - I am unsure if I am legally allowed to redistribute their data... if it's a problem I will ask for forgiveness later
    - See FAQ for how to get access yourself
- Windows? You could probably get this to work in another OS but I'm not testing it.

## How to use the program:
1. Make sure PowerShell 7 is installed, see requirements. You should have an app called "PowerShell 7" searchable.
2. Download the GitHub repo as a zip file, and extract it. Run the program by either:
    - Run run.bat via File Explorer. This should set correct ExecutionPolicy and load the GUI after a few seconds.
    - Run .\iracing-calendar-spreadsheet.ps1 via PowerShell 7 or similar Terminal. You may need to use Set-ExecutionPolicy if you are unfamiliar with running unsigned powershell scripts. See Appendix.
3. Choose JSON file if it does not already select it, likely in the jsons folder.
4. Click items in the left list box to see the tracks in that series.
5. Click the checkbox to add that series to your export/calendar.
6. Repeat as needed. 8 series are recommened for the small template. The background turns green when you reach 8.
7. If you'd prefer to output all the series, click that checkbox. In the XLSX file below, use the sheet "iRacingCalendarTemplateLarge".
8. Choose output filename and path or leave the default.
9. Click Create CSV.

## How to create the spreadsheet:
1. Open the created CSV in Excel.
2. Column B will be all of the track series information. There is a formatting bug and it does not properly format the CSV. Select Column B from row 1 to row 16
3. Use the "Text to Columns" feature in Excel to format that column properly. To find it, Data > Data Tools section > Text to Columns. If you are unfamiliar, see "Text to Columns" in appendix. 
4. Open the template file you want to use. Currently there is just one, "TemplateAll.xlsx".
5. As of March 2025, there are two tabs/sheets, one is the intended small template for 8 series and one for all the series.
6. Once the CSV is properly formatted from step 3, copy the columns you want in the template. 8 series recommended for small template.
7. Click back into the Excel template and click cell B1.
8. Use the "Paste Special" feature and choose to copy "Values". The keyboard shortcut Ctrl+Shift+V may work. The formatting/cell borders should stay the same.
9. Update the first "Week Start" cell (J5) to the start of the current season.
10. Print! If your landscape 8.5x11 page does not fit 8 series of 12 weeks, you may need to adjust column/row sizes, reduce font size, or lower page margins.

## Known Issues

1. iRacing's "time" values don't all use the same formatting, so some series may not be minimized to fit the time cell. The first row is set to "shrink" values and most series should be readable.
2. Series that run more than 12 weeks or less than 12 weeks are skipped. At the moment I have no interest in keeping track of them myself, and they complicate the table if we need to handle them as edge cases.
3. Series list does not refresh when choosing a new series json. May need to rename old one or change default in script.

## Possible Improvements

1. Calculate "Week Start" and "Week End" from json data and include in the CSV.
2. Export direct to formatted Excel (was original plan, hard to do).
3. In my manual version, I bolded tracks that I have, and counted them at the ends of each row/column. Might be possible, especially if #2 was done.
4. Rebuild in another language (javascript for web-based usage?) or in PowerShell 5 which might be better supported (and support ps2toexe)?

## FAQ

Q: There are already websites like schedule4i.racing and irbg.net for this, why did you make this?

A: I started by making a spreadsheet template of the series I wanted to run and printing it next to my computer, usually using these tools and editing the data to fit the template. Knowing enough about coding and the iRacing API to be dangerous (and a helpful suggestion from iRacing Support), I realized I could automate some of the dirty work.

Q: Why PowerShell?

A: I work in Windows tech support and was already familiar with creating and editing PowerShell scripts. A lot of this work was assisted by Copilot or ChatGPT.

Q: Why PowerShell 7?

A: There is a feature of Export-CSV (-NoHeader) in PowerShell 7 that works better for the intended purpose.

Q: You should use X language.

A: I will consider refactoring this in another language once this is released and working. For now, I'm just happy to be able to make my own calendars.

Q: How do I download my own json file?

A: I used the documentation at https://www.postman.com/rankupgamers/iracing-new-api/documentation/uc5dzd8/iracing . You must first turn on legacy authentication at https://oauth.iracing.com/accountmanagement/security/ then login and follow the following iRacing link: https://members-login.iracing.com/?ref=https://members-ng.iracing.com/data/series/seasons. 

## Appendix

### How to bypass ExecutionPolicy errors

If you get an error containing "cannot be loaded because running scripts is disabled on this system", you need to use Set-ExecutionPolicy -Scope <scope> -ExecutionPolicy <executionpolicy> to change the security polices for running PowerShell scripts. Changing this can make your computer less secure, but a simple way to fix this is to change the Process Scope and then run the ps1 script, for example:

        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
        .\iracing-calendar-spreadsheet.ps1

If you are concerned with your computer's security, run Get-ExecutionPolicy and verify all the polices are set to Undefined. From Microsoft: "[If no execution policy is set in any scope, the effective execution policy is Restricted, which is the default for Windows clients.](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.5#managing-the-execution-policy-with-powershell)"

### How to use "Text to Columns"
Basic steps: Step 1, choose delimited, click Next. Step 2, add Comma as a delimiter, you should see data preview separate into columns, click next. Step 3, can ignore this screen, click Finish.

See link for a video guide from MS:
https://support.microsoft.com/en-us/office/split-text-into-different-columns-with-the-convert-text-to-columns-wizard-30b14928-5550-41f5-97ca-7a3e9c363ed7


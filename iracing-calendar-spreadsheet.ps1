# iRacing Calendar Spreadsheet Creator
# Takes existing series json from the iRacing web API and converts it into a table prepared for excel sheet template

# Variable to set the default json and output file to use every season
$defaultJSONPath = ".\jsons\season-series-25s2.json"
$defaultOutputPath = ".\outputs\output.csv"
function trackNameMinimizer ($trackName) {
    # Take a track name and run a variety of replace functionality to minimize the name

    # Define the replacement map as an array of key-value pairs
    $replacements = @(

        # Specific track replacements
        @{ Key = "Circuit des 24 Heures du Mans"; Value = "Le Mans" }
        @{ Key = "Virginia International Raceway"; Value = "VIR" }
        @{ Key = "Autodromo Internazionale Enzo e Dino Ferrari"; Value = "Imola" }
        @{ Key = "Nürburgring Nordschleife"; Value = "Nordschleife" }
        @{ Key = "Nürburgring"; Value = "Nurburg" }
        @{ Key = "Hockenheimring Baden-Württemberg"; Value = "Hockenheim" }
        @{ Key = "Circuit de Lédenon"; Value = "Ledenon" }
        @{ Key = "WeatherTech Raceway at Laguna Seca"; Value = "Laguna Seca" }
        @{ Key = "Autodromo Internazionale del Mugello"; Value = "Mugello" }
        @{ Key = "Circuit de Barcelona Catalunya"; Value = "Barcelona" }
        @{ Key = "Autodromo Nazionale Monza"; Value = "Monza" }
        @{ Key = "Misano World Circuit Marco Simoncelli"; Value = "Misano Sic" }
        @{ Key = "Circuit de Spa-Francorchamps"; Value = "Spa" }
        @{ Key = "Autódromo José Carlos Pace"; Value = "Interlagos" }
        @{ Key = "Long Beach Street Circuit"; Value = "Long Beach" }
        @{ Key = "Canadian Tire Motorsports Park"; Value = "Mosport" }
        @{ Key = "Detroit Grand Prix at Belle Isle"; Value = "Detroit Belle Isle" }
        @{ Key = "Mobility Resort Motegi"; Value = "Motegi" }
        @{ Key = "Circuit of the Americas"; Value = "COTA" }
        @{ Key = "World Wide Technology Raceway (Gateway)"; Value = "Gateway" }
        @{ Key = "Circuit de Jerez - Ángel Nieto"; Value = "Jerez" }
        @{ Key = "Lucas Oil Indianapolis Raceway Park"; Value = "IRP" }
        @{ Key = "Daytona Rallycross and Dirt Road"; Value = "Daytona" }
        @{ Key = "Kevin Harvick's Kern Raceway"; Value = "Harvick's Kern" }
        @{ Key = "Federated Auto Parts Raceway at I-55"; Value = "I-55" }
        @{ Key = "Lånkebanen (Hell RX)"; Value = "Hell RX" }
        @{ Key = "MotorLand Aragón"; Value = "Aragon" }

        # General wording replacements
        # Some of these are probably superfluous with all the above tracks
        @{ Key = " International Circuit"; Value = "" }
        @{ Key = " Racing Circuit"; Value = "" }
        @{ Key = " Motorsenter"; Value = "" }
        @{ Key = " International Raceway"; Value = "" }
        @{ Key = " International Speedway"; Value = "" }
        @{ Key = " International Racing Course"; Value = "" }
        @{ Key = " Motor Raceway"; Value = "" }
        @{ Key = " Motor Speedway"; Value = "" }
        @{ Key = " International"; Value = "" }
        @{ Key = " Superspeedway"; Value = "" }
        @{ Key = " Motorsports Park"; Value = "" }
        @{ Key = "Circuit de "; Value = "" }
        @{ Key = "Circuito de "; Value = "" }
        @{ Key = "Circuit "; Value = "" }
        @{ Key = " Circuit"; Value = "" }
        @{ Key = "Motorsport Arena "; Value = "" }
        @{ Key = " Speedway"; Value = "" }
        @{ Key = " Sports Car Course"; Value = "" }
        @{ Key = " Street Circuit"; Value = "" }
        @{ Key = "[Legacy]"; Value = "[L]" }
    )

    # Iterate through the replacements in order
    foreach ($replacement in $replacements) {
        $trackName = $trackName -replace [regex]::Escape($replacement.Key), $replacement.Value
    }

    return $trackName
}

function trackConfigMinimizer ($trackConfig) {
    # Take a track config and run a variety of replace functionality to minimize the name

    # Define the replacement map
    # This is done in the original hashtable method, could be changed later
    $replacements = @{
        "International" = "Intl"
        "Grand Prix"   = "GP"
        "Road Course" = "RC"
        "Summit Point Raceway" = ""
        "Full Course" = "Full"
        " Circuit" = ""
        "24 Heures du Mans" = ""
        "Industriefahrten" = ""
        "Belle Isle" = ""
    }
    
    # For the track config input, perform replacements from above map
    foreach ($key in $replacements.Keys) {
        $trackConfig = $trackConfig -replace [regex]::Escape($key), $replacements[$key]
    }

    Write-Output $trackConfig

}

function scheduleTimeMinimizer ($scheduleTime) {
    # Takes the "schedule description" variable which includes the time of each series and minimizes it
    # This can be improved, perhaps to adjust for one's local time vs. GMT?

    # Define the replacement map
    $replacements = @(

        # This should get the weird Ring Meister time to be more readable
        @{ Key = " `| Qualifying every even 2 hours at :30"; Value = "" }

        # Regular replacements to reduce Time field
        @{ Key = "Races "; Value = "" }
        @{ Key = "every hour"; Value = "hourly" }
        @{ Key = " past"; Value = "" }
        @{ Key = "minutes"; Value = "mins" }
        @{ Key = "at :00 and :30"; Value = ": :00`/:30" }
        @{ Key = "at :15 and :45"; Value = ": :00`/:30" }

    )
    
    # Perform replacements
    foreach ($replacement in $replacements) {
        $scheduleTime = $scheduleTime -replace [regex]::Escape($replacement.Key), $replacement.Value
    }

    return $scheduleTime

}

function getProperTrackName ($trackObject) {
    # Get the "proper" track name from the track object
    # Proper as in it is fully minimized for the spreadsheet
    # This uses trackNameMinimizer and trackConfigMinimizer and combines that process

    # use trackNameMinimizer on long track name
    $trackBase = trackNameMinimizer($trackObject.track.track_name)
    $trackFinalName = ""

    # if track has a config name, add it to track name, else use trackNameMinimizer output
    if ($track.track.config_name) { 
        $trackConfigShort = trackConfigMinimizer($track.track.config_name)
        $trackFinalName = $trackBase + " " + $trackConfigShort
    }
    else {
        $trackFinalName = $trackBase
    }
    
    return $trackFinalName
}

function specialSeriesConfig ($track) {

    # If track name contains Ringmeister, return the cars instead of the track
    # We are being lazy and only grabbing the 1st car to fit the cell
    # Future code could check for GT3, GT4, etc. and return that.
    if ($track.season_name -like "*Ring Meister*") {

        $car = $track.race_week_cars[0].car_name

        return $car
    }
    elseif ($track.season_name -like "*Draft Master*") {
        # For Draft Masters, return the track and cars.

        $trackName = getProperTrackName($track)
        $carName = $track.race_week_cars[0].car_name
    
        return $trackName + " with " + $carName

    }

}

function populateSeriesListBox () {
    # Populate the seriesListBox with the series names from the JSON data

    foreach ($series in $jsonData) {

        # If series is 12 weeks, add. Need to fix later if we want to include smaller series.
        if ($series.schedules.Count -eq 12) {
        
            [void] $seriesListBox.Items.Add($series.season_name)

        }
        # For now, if not 12 weeks just add a note that it is not supported
        else {
            [void] $seriesListBox.Items.Add("NOT SUPPORTED "+$series.season_name)
        }

    }
}

function GUI () {
    # Program GUI that is called later in the script.
    # This function creates a GUI that allows the user to select series and generate a CSV file.

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'iRacing Calendar Spreadsheet Creator'
    $form.Size = New-Object System.Drawing.Size(900,550)
    $form.StartPosition = 'CenterScreen'

    # Textbox for JSON file path
    $jsonFilePathLabel = New-Object System.Windows.Forms.Label
    $jsonFilePathLabel.Location = New-Object System.Drawing.Point(10,20)
    $jsonFilePathLabel.Size = New-Object System.Drawing.Size(200,20)
    $jsonFilePathLabel.Text = 'JSON file path:'
    $form.Controls.Add($jsonFilePathLabel)

    $jsonFilePathTextBox = New-Object System.Windows.Forms.TextBox
    $jsonFilePathTextBox.Location = New-Object System.Drawing.Point(10,40)
    $jsonFilePathTextBox.Size = New-Object System.Drawing.Size(300,20)
    $jsonFilePathTextBox.Text = "Choose JSON file >"
    $form.Controls.Add($jsonFilePathTextBox)

    # Button to browse for JSON file
    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(320,40)
    $browseButton.Size = New-Object System.Drawing.Size(75,23)
    $browseButton.Text = 'Browse'
    $form.Controls.Add($browseButton)
    $browseButton.Add_Click({
        # When clicking the browse button, open a file dialog to select the JSON file

        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $openFileDialog.Title = "Select JSON File"

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $jsonFilePathTextBox.Text = $openFileDialog.FileName

            # Read and convert the JSON file into a PowerShell object
            # Trying global variable because problems with future usage.
            $global:jsonData = Get-Content $jsonFilePathTextBox.Text | ConvertFrom-Json

            $seriesListBox.Items.Clear()

            # Add series names to the list box
            populateSeriesListBox

        }
    })

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,80)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Please select your series:'
    $form.Controls.Add($label)

    # Checked list box for iRacing Series
    $seriesListBox = New-Object System.Windows.Forms.CheckedListBox
    $seriesListBox.Location = New-Object System.Drawing.Point(10,100)
    $seriesListBox.Size = New-Object System.Drawing.Size(260,360)
    $form.Controls.Add($seriesListBox)

    # If a default JSON file is set, populate the series list box, to save time
    if (Test-Path $defaultJSONPath) {
        $jsonFilePathTextBox.Text = $defaultJSONPath
        $jsonData = Get-Content $defaultJSONPath | ConvertFrom-Json
        populateSeriesListBox
    }

    $trackLabel = New-Object System.Windows.Forms.Label
    $trackLabel.Location = New-Object System.Drawing.Point(300,80)
    $trackLabel.Size = New-Object System.Drawing.Size(280,20)
    $trackLabel.Text = 'Selected series uses:'
    $form.Controls.Add($trackLabel)

    $trackTextBox = New-Object System.Windows.Forms.TextBox
    $trackTextBox.Location = New-Object System.Drawing.Point(300,100)
    $trackTextBox.Size = New-Object System.Drawing.Size(500,250)
    $trackTextBox.Multiline = $true
    $trackTextBox.ScrollBars = 'Vertical'
    $form.Controls.Add($trackTextBox)

    # Single checkbox to generate csv for all series
    $allSeriesCheckbox = New-Object System.Windows.Forms.CheckBox
    $allSeriesCheckbox.Location = New-Object System.Drawing.Point(300,360)
    $allSeriesCheckbox.Size = New-Object System.Drawing.Size(200,20)
    $allSeriesCheckbox.Text = 'Generate CSV for all series'
    $form.Controls.Add($allSeriesCheckbox)

    # Textbox for output file path
    $outputFilePathLabel = New-Object System.Windows.Forms.Label
    $outputFilePathLabel.Location = New-Object System.Drawing.Point(300,390)
    $outputFilePathLabel.Size = New-Object System.Drawing.Size(200,20)
    $outputFilePathLabel.Text = 'Output file path:'
    $form.Controls.Add($outputFilePathLabel)
    $outputFilePathTextBox = New-Object System.Windows.Forms.TextBox
    $outputFilePathTextBox.Location = New-Object System.Drawing.Point(300,410)
    $outputFilePathTextBox.Size = New-Object System.Drawing.Size(300,20)
    if (Test-Path $defaultOutputPath) {
        $outputFilePathTextBox.Text = $defaultOutputPath
    }
    else {
        $outputFilePathTextBox.Text = "Select output file >"
    }
    $form.Controls.Add($outputFilePathTextBox)

    # Button to browse for output file path
    $outputBrowseButton = New-Object System.Windows.Forms.Button
    $outputBrowseButton.Location = New-Object System.Drawing.Point(610,410)
    $outputBrowseButton.Size = New-Object System.Drawing.Size(75,23)
    $outputBrowseButton.Text = 'Browse'
    $form.Controls.Add($outputBrowseButton)
    $outputBrowseButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveFileDialog.Title = "Save CSV File"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputFilePathTextBox.Text = $saveFileDialog.FileName
        }
    })

    # create CSV button
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Location = New-Object System.Drawing.Point(75,470)
    $createButton.Size = New-Object System.Drawing.Size(75,23)
    $createButton.Text = 'Create CSV'
    $form.Controls.Add($createButton)

    # close Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150,470)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Close'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    # Label to count number of series checked, since the goal is to get 8
    $numberSeriesChecked = 0

    $seriesCheckedCountLabel = New-Object System.Windows.Forms.Label
    $seriesCheckedCountLabel.Location = New-Object System.Drawing.Point(500,40)
    $seriesCheckedCountLabel.Size = New-Object System.Drawing.Size(280,20)
    $seriesCheckedCountLabel.Text = 'Number of checked series: '+$numberSeriesChecked
    $form.Controls.Add($seriesCheckedCountLabel)

    # On item check, update $numberSeriesChecked
    # When 8, turn green, if over, reset color
    # This isn't perfect, in testing you can make the array lose track of what is checked?
    $seriesListBox.add_ItemCheck({
        $numberSeriesChecked = $seriesListBox.CheckedItems.Count+1
        $seriesCheckedCountLabel.Text = 'Number of checked series: '+$numberSeriesChecked

        if ($numberSeriesChecked -eq 8) {
            $form.BackColor = "Green"
        }
        if ($numberSeriesChecked -gt 8) {
            $form.BackColor = ""
        }
    })

    $form.Topmost = $true

    # When clicking the create button, add checked items to array, then run createSeriesCSV
    $createButton.Add_Click({

        # Use the checkedItems variable
        $checkedItems = @()
        $seriesToExport = @()

        # If the allSeriesCheckbox is checked, just put all series into checkedItems
        if ($allSeriesCheckbox.Checked) {
            $checkedItems = $seriesListBox.Items
        }
        # Otherwise, only put the checked items into checkedItems
        else {
        
            $checkedItems = $seriesListBox.CheckedItems
        }

        # Not really sure if this is needed?? But lets not break it while it's working.
        foreach ($item in $checkedItems) {
            $seriesToExport += $item
        }
        createSeriesCSV($seriesToExport)
    })


    # When selecting an item on the Series listBox on left, add the tracks to the right hand trackTextBox
    $seriesListBox.add_SelectedIndexChanged({
        
        $selectedSeries = $jsonData | Where-Object { $_.season_name -eq $seriesListBox.SelectedItem }
        
        $trackTextBox.Clear()
        # Error checking for the $jsonData variable
        if ($null -eq $selectedSeries) {
            $trackTextBox.AppendText("Error selecting series.`r`n")
        }
        # Ring Meister series
        elseif ($selectedSeries.season_name -like "*Ring Meister*") {
            $trackTextBox.AppendText("Cars:`r`n")
            foreach ($track in $selectedSeries.schedules) {
                $trackName = specialSeriesConfig($track)
                $trackTextBox.AppendText($trackName + "`r`n")
            }
        }
        # Draft Masters series
        elseif ($selectedSeries.season_name -like "*Draft Master*") {
            $trackTextBox.AppendText("Tracks with Cars:`r`n")
            foreach ($track in $selectedSeries.schedules) {
                $trackName = specialSeriesConfig($track)
                $trackTextBox.AppendText($trackName + "`r`n")
            }
        }
        # Regular series
        else {
            $trackTextBox.AppendText("Tracks:`r`n")
            foreach ($track in $selectedSeries.schedules) {
                $trackName = getProperTrackName($track)
                $trackTextBox.AppendText($trackName + "`r`n")
            }
        }
    })

    

    $form.ShowDialog()
}

function createSeriesCSV ($selectedSeries) {
    # Function that creates the CSV file

    # Initialize arrays for rows in the CSV
    $rowID = @()
    $rowTime = @()
    $rowLicense = @()
    $rowStyle = @()
    $rowName = @()
    $rowT1 = @()
    $rowT2 = @()
    $rowT3 = @()
    $rowT4 = @()
    $rowT5 = @()
    $rowT6 = @()
    $rowT7 = @()
    $rowT8 = @()
    $rowT9 = @()
    $rowT10 = @()
    $rowT11 = @()
    $rowT12 = @()

    # For each series in the JSON data
    foreach ($series in $jsonData) {

        # if the series is in the selectedSeries array, add it to the CSV
        if ($series.season_name -in $selectedSeries) {

            # If series is 12 weeks, get info. Need to fix later if we want to include smaller series.
            if ($series.schedules.Count -eq 12) {

                # Get series info
                $rowTime += scheduleTimeMinimizer($series.schedule_description)
                $rowLicenseNum = $series.license_group
                switch ($rowLicenseNum) {
                    1 {$rowLicense += "Rookie"}
                    2 {$rowLicense += "D"}
                    3 {$rowLicense += "C"}
                    4 {$rowLicense += "B"}
                    5 {$rowLicense += "A"}
                }
                $rowName += ($series.season_name -split " -")[0]
                $rowStyle += $series.track_types.track_type

                # not used yet, probably can remove this and the array
                $rowID += $series.season_id

                # iterator for for loop
                $trackNum = 0
                # iterate on track list
                foreach ($track in $series.schedules) {

                    # Special case for Ring Meister series, return cars instead of tracks
                    # Can combine Ring Meister and Draft Masters... not sure why I separated them
                    if ($series.season_name -like "*Ring Meister*") {
                        $trackFull = specialSeriesConfig($track)
                    }
                    # Add elseif for Draft Masters
                    elseif ($series.season_name -like "*Draft Master*") {
                        $trackFull = specialSeriesConfig($track)
                    }
                    else {
                        $trackFull = getProperTrackName($track)
                    }

                    # Assign $trackFull to the appropriate $rowT variable
                    switch ($trackNum) {
                        0 { $rowT1 += $trackFull }
                        1 { $rowT2 += $trackFull }
                        2 { $rowT3 += $trackFull }
                        3 { $rowT4 += $trackFull }
                        4 { $rowT5 += $trackFull }
                        5 { $rowT6 += $trackFull }
                        6 { $rowT7 += $trackFull }
                        7 { $rowT8 += $trackFull }
                        8 { $rowT9 += $trackFull }
                        9 { $rowT10 += $trackFull }
                        10 { $rowT11 += $trackFull }
                        11 { $rowT12 += $trackFull }
                    }

                    # Increment iterator
                    $trackNum++
                }
            }
            else {
                # Series is less than 12 weeks, ignore for now.
                # For series of 6 weeks, we could just alternate the rows we put the track into?
            }

        }
    }

    # Build the spreadsheet/csv/clipboard item to import into Excel of all of the rows in order.

    # Prepare data for export
    $outputData = @(
        [PSCustomObject]@{ RowType = "Time"; Data = ($rowTime -join ",") }
        [PSCustomObject]@{ RowType = "License"; Data = ($rowLicense -join ",") }
        [PSCustomObject]@{ RowType = "Style"; Data = ($rowStyle -join ",") }
        [PSCustomObject]@{ RowType = "Name"; Data = ($rowName -join ",") }
        [PSCustomObject]@{ RowType = "Track1"; Data = ($rowT1 -join ",") }
        [PSCustomObject]@{ RowType = "Track2"; Data = ($rowT2 -join ",") }
        [PSCustomObject]@{ RowType = "Track3"; Data = ($rowT3 -join ",") }
        [PSCustomObject]@{ RowType = "Track4"; Data = ($rowT4 -join ",") }
        [PSCustomObject]@{ RowType = "Track5"; Data = ($rowT5 -join ",") }
        [PSCustomObject]@{ RowType = "Track6"; Data = ($rowT6 -join ",") }
        [PSCustomObject]@{ RowType = "Track7"; Data = ($rowT7 -join ",") }
        [PSCustomObject]@{ RowType = "Track8"; Data = ($rowT8 -join ",") }
        [PSCustomObject]@{ RowType = "Track9"; Data = ($rowT9 -join ",") }
        [PSCustomObject]@{ RowType = "Track10"; Data = ($rowT10 -join ",") }
        [PSCustomObject]@{ RowType = "Track11"; Data = ($rowT11 -join ",") }
        [PSCustomObject]@{ RowType = "Track12"; Data = ($rowT12 -join ",") }
    )

    # Export to CSV
    $outputFile = $outputFilePathTextBox.Text
    $outputData | Export-Csv -Path $outputFile -NoHeader -NoTypeInformation -Encoding UTF8

    $messageBoxTitle = "Export Complete"
    $messageBoxMessage = "Data exported to $outputFile"
    [System.Windows.Forms.MessageBox]::Show($messageBoxMessage, $messageBoxTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

}

# run GUI
# GUI should then run the createSeriesCSV function which makes the magic
# justPowerShellThings
GUI
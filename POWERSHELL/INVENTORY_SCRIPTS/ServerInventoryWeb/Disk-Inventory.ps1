# Load Form Assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Begin drawing forms
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Get disk inventory of specified computer."
$Form.Size = New-Object System.Drawing.Size(800,400)
$Form.StartPosition = "CenterScreen"
$Form.KeyPreview = $True
$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter") {
   &$DiskInventory
}})
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape") {
    $Form.Close()
}})

# Create the label for the textbox to enter Computer Name in
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(5,26)
$Label.Size = New-Object System.Drawing.Size (90,30)
$Label.Text = "Computer Name"
$Form.Controls.Add($Label)

# Create the Textbox to enter the computer name into
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Size(100,20)
$TextBox.Size = New-Object System.Drawing.Size(200,30)
$TextBox.Text = 'Enter computer name here.'
$Form.Controls.Add($TextBox)

# Create status bar for script progress
$StatusBar1 = New-Object System.Windows.Forms.StatusBar
$StatusBar1.Name = "StatusBar1"
$StatusBar1.Text = "Ready..."
$Form.Controls.Add($StatusBar1)

# Run script Get-DiskInventory on QueryButton click
$DiskInventory = {
    $StatusBar1.Text = "Retreiving disk data for $($TextBox.Text)..."

    $ResultsTextBox.Text = foreach-object  {
        Get-WmiObject -Class win32_logicaldisk -ComputerName $TextBox.Text -filter "drivetype='3'"
    } |
    Sort-Object -Property SystemName,DeviceID |
    Select-Object -Property SystemName,DeviceID,VolumeName,
    @{Label='Size(GB)';Expression={$_.Size / 1GB -as [int]}},
    @{Label='Freespace(GB)';Expression={$_.Freespace / 1GB -as [int]}},
    @{Label='% FreeSpace';Expression={$_.FreeSpace / $_.Size * 100 -as [int]}} |
    Format-List |
    Out-String

    $StatusBar1.Text = "Testing Complete"
}

# Create Button to launch script
$QueryButton = New-Object System.Windows.Forms.Button
$QueryButton.Location = New-Object System.Drawing.Size(310,16)
$QueryButton.Size = New-Object System.Drawing.Size(70,24)
$QueryButton.Text = "Query"
$QueryButton.Add_Click($DiskInventory)
$Form.Controls.Add($QueryButton)

# Create results display label
$ResultsLabel = New-Object System.Windows.Forms.Label
$ResultsLabel.Location = New-Object System.Drawing.Size(5,60)
$ResultsLabel.Size = New-Object System.Drawing.Size(180,16)
$ResultsLabel.Text = "Get-DiskInventory Results"
$Form.Controls.Add($ResultsLabel)

# Create results textbox to display the results of Get-DiskInventory
$ResultsTextBox = New-Object System.Windows.Forms.TextBox
$ResultsTextBox.Location = New-Object System.Drawing.Size(5,86)
$ResultsTextBox.Size = New-Object System.Drawing.Size(200,800)
$ResultsTextBox.Text = "Waiting for info..."
$Form.Controls.Add($ResultsTextBox)

# Show Form
$Form.TopMost = $True
$Form.Add_Shown({$Form.Activate})
[void] $Form.ShowDialog()
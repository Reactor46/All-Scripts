Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.ServiceModel

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "App Pool Recycler"
$form.Size = New-Object System.Drawing.Size(400,300)
$form.StartPosition = "CenterScreen"

# Create labels
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(10, 10)
$serverLabel.Size = New-Object System.Drawing.Size(200, 20)
$serverLabel.Text = "Select Server:"
$form.Controls.Add($serverLabel)

$appPoolLabel = New-Object System.Windows.Forms.Label
$appPoolLabel.Location = New-Object System.Drawing.Point(10, 80)
$appPoolLabel.Size = New-Object System.Drawing.Size(150, 20)
$appPoolLabel.Text = "Select App Pool:"
$form.Controls.Add($appPoolLabel)

# Create ComboBoxes
$serverCombo = New-Object System.Windows.Forms.ComboBox
$serverCombo.Location = New-Object System.Drawing.Point(10, 30)
$serverCombo.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($serverCombo)

$appPoolCombo = New-Object System.Windows.Forms.ComboBox
$appPoolCombo.Location = New-Object System.Drawing.Point(10, 100)
$appPoolCombo.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($appPoolCombo)

# Query App Pools button
$queryButton = New-Object System.Windows.Forms.Button
$queryButton.Location = New-Object System.Drawing.Point(10, 150)
$queryButton.Size = New-Object System.Drawing.Size(150, 30)
$queryButton.Text = "Query App Pools"
$queryButton.Add_Click({
    $selectedServer = $serverCombo.SelectedItem
    if ($selectedServer) {
        Query-AppPools -Server $selectedServer
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a server first.")
    }
})
$form.Controls.Add($queryButton)

# Recycle App Pool button
$recycleButton = New-Object System.Windows.Forms.Button
$recycleButton.Location = New-Object System.Drawing.Point(10, 200)
$recycleButton.Size = New-Object System.Drawing.Size(150, 30)
$recycleButton.Text = "Recycle App Pool"
$recycleButton.Add_Click({
    $selectedServer = $serverCombo.SelectedItem
    $selectedAppPool = $appPoolCombo.SelectedItem
    if ($selectedServer -and $selectedAppPool) {
        Invoke-Command -ComputerName $selectedServer -ScriptBlock {
            Import-Module WebAdministration
            Restart-WebAppPool -Name $using:selectedAppPool
        }
        [System.Windows.Forms.MessageBox]::Show("App Pool '$selectedAppPool' on server '$selectedServer' has been recycled.")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a server and an application pool first.")
    }
})
$form.Controls.Add($recycleButton)

# Function to query application pools on the selected server

function Query-AppPools {
    param (
        [string]$Server
    )
    $appPools = Invoke-Command -ComputerName $Server -ScriptBlock {
        Import-Module WebAdministration
        Get-ChildItem IIS:\AppPools | Select-Object -ExpandProperty Name
    }
    $appPoolCombo.Items.Clear()
    $appPoolCombo.Items.AddRange($appPools)
}

# Populate server list
$servers = "FBV-SCRCD10-P01", "FBV-SCRCD10-P02", "FBV-SCRCD10-P03","FBV-SCRCD10-P04"  # Add your server names here
$serverCombo.Items.AddRange($servers)

# Show the form
$form.ShowDialog() | Out-Null

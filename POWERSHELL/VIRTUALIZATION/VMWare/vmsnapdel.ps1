#
# VMware Snapshot Delete Utility
#

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()

Import-Module -Name VMware.VimAutomation.Core

#-------------------------------------------------------------------------------

# Connection status
$Script:connected = $false

# Connect/disconnect from VI server
$connect_to_VIServer = { 
    $VIServer = $tbox_VIServer.Text.ToString()
    if ($Script:connected) {
        Disconnect-VIServer -Confirm:$false
        $btn_VIServer.Text = "Connect"
        $lbl_ConnectedTo.Text = "Disconnected"
        $Script:connected = $false
        $grid_Snapshots.DataSource = $null
    }
    else {
        if (Connect-VIServer $VIServer) {
            $lbl_ConnectedTo.Text = "Connected to $VIServer"
            $btn_VIServer.Text = "Disconnect"
            $Script:connected = $true
            $snapshots_list = [system.Array] (
                Get-VM | 
                Get-Snapshot |
                Select-Object VM, Name, SizeGB, Created, Description |
                Sort-Object -Property Created
            )
            $grid_Snapshots.DataSource = [system.Collections.ArrayList] $snapshots_list
        }
    }
}

# Delete selected snapshot
$delete_selected = {
    if ([System.Windows.Forms.MessageBox]::Show("Do you want to delete the selected snapshots?",
            "Delete Snapshots", [System.Windows.Forms.MessageBoxButtons]::OKCancel) -eq "OK") {
        $progressBar.Maximum = $grid_Snapshots.SelectedRows.Count
        $progressBar.Minimum = 0

        $btn_VIServer.Enabled = $false
        $btn_DeleteSelected.Enabled = $false

        $grid_Snapshots.SelectedRows |
        ForEach-Object {
            $snapshot = Get-Snapshot -VM $_.Cells[0].Value
            Remove-Snapshot -Snapshot $snapshot -Confirm:$false
            $progressBar.Value++
        }
        Start-Sleep -Seconds 2
        $progressBar.Value = 0
        $snapshots_list = [system.Array] (
            Get-VM | 
            Get-Snapshot |
            Select-Object VM, Name, SizeGB, Created, Description |
            Sort-Object -Property Created
        )
        $grid_Snapshots.DataSource = [system.Collections.ArrayList] $snapshots_list
        
        $btn_VIServer.Enabled = $true
        $btn_DeleteSelected.Enabled = $true
    }
}

#-------------------------------------------------------------------------------

# Main Form
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = 'VMware Snapshot Delete Utility'
$main_form.Width = 800
$main_form.Height = 600
$main_form.Icon = New-Object System.Drawing.Icon ( $PSScriptRoot.ToString() + "\vmware.ico" )

# Label
$lbl_VIServer = New-Object System.Windows.Forms.Label
$lbl_VIServer.Text = "VIServer"
$lbl_VIServer.Location = New-Object System.Drawing.Point(10, 13)
$lbl_VIServer.AutoSize = $true
$main_form.Controls.Add($lbl_VIServer)

# This label change with connection status "Disconnected" or "Connected to"
$lbl_ConnectedTo = New-Object System.Windows.Forms.Label
$lbl_ConnectedTo.Location = New-Object System.Drawing.Point(10, 33)
$lbl_ConnectedTo.Text = "Disconnected"
$lbl_ConnectedTo.AutoSize = $true
$main_form.Controls.Add($lbl_ConnectedTo)

# Input text box. It takes the IP address / hostnmame of the VI server
$tbox_VIServer = New-Object System.Windows.Forms.TextBox
$tbox_VIServer.Location = New-Object System.Drawing.Point(70, 10)
$tbox_VIServer.Size = New-Object System.Drawing.Size(260, 20)
$tbox_VIServer.Text = ""
$main_form.Controls.Add($tbox_VIServer)

# Connect/disconnect button
$btn_VIServer = New-Object System.Windows.Forms.Button
$btn_VIServer.Text = "Connect"
$btn_VIServer.Location = New-Object System.Drawing.Point(350, 9)
$btn_VIServer.AutoSize = $true
$btn_VIServer.Add_Click($connect_to_VIServer)
$main_form.Controls.Add($btn_VIServer)

# Snapshots data grid
$grid_Snapshots = New-Object System.Windows.Forms.DataGridView
$grid_Snapshots.Location = New-Object System.Drawing.Point(10, 60)
$grid_Snapshots.Size = New-Object System.Drawing.Size(760, 460)
$grid_Snapshots.ColumnHeadersVisible = $true
$grid_Snapshots.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid_Snapshots.AutoSizeColumnsMode = 'AllCells'
$grid_Snapshots.AutoSize = $true
$grid_Snapshots.Margin = '4, 4, 4, 4'
$grid_Snapshots.TabIndex = 0
$grid_Snapshots.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left -bor
    [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Top
)
$main_form.Controls.Add($grid_Snapshots)

# Delete the snapshots selected in the data grid
$btn_DeleteSelected = New-Object System.Windows.Forms.Button
$btn_DeleteSelected.Text = "Delete selected"
$btn_DeleteSelected.Location = New-Object System.Drawing.Point(680, 530)
$btn_DeleteSelected.AutoSize = $true
$btn_DeleteSelected.Add_Click($delete_selected)
$btn_DeleteSelected.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right
)
$main_form.Controls.Add($btn_DeleteSelected)

# Progress label
$lbl_progress = New-Object System.Windows.Forms.Label
$lbl_progress.Text = "progress"
$lbl_progress.Location = New-Object System.Drawing.Point(10, 530)
$lbl_progress.AutoSize = $true
$lbl_progress.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left
)
$main_form.Controls.Add($lbl_progress)

# Porgress bar to display deletion progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(65, 533)
$progressBar.Size = New-Object System.Drawing.Size(400, 10)
$progressBar.Style = "Blocks"
$progressBar.Value = 0
$progressBar.Anchor = (
    [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Left
)
$main_form.Controls.Add($progressBar)


[void] $main_form.ShowDialog()

#EOF#
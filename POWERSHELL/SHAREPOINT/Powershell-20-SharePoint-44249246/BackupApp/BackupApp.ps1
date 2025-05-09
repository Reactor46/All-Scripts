$ErrorActionPreference = "stop"
$DebugMode=0

# Load external assemblies
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
if (!(Get-PSSnapin|Where-Object {$_.Name -like '*Sharepoint*'})) {
    write-host "Adding PSSnapin Microsoft.SharePoint.PowerShell"
    try {
        Add-PSSnapin "Microsoft.SharePoint.PowerShell"
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show("SharePoint Snapin not available, are you on a SharePoint server?")
        break
    }
    finally {}
}

if (!(Get-Module|Where-Object {$_.Name -like '*webadministration*'})) {
    write-host "Importing module webadministration"
    try {
        import-module webadministration -ea continue
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show("Webadministration module not available, are you on an IIS server?")
        break
    }
    finally {}
}
$backedupsites = @{}
$Form = new-object System.Windows.Forms.form
$Form.ClientSize = new-object System.Drawing.Size(400, 520)
$MainToolbar = new-object System.Windows.Forms.MenuStrip
$fileToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$exitToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$resetToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$debugToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem

function updateDebugLabels {
	if ($global:DebugMode) {
        foreach ($site in $global:backedupsites) {$global:BackedUpSiteList = ""
        } 
        $global:DebugLabelText.Text = "Debug info will show up here:" + 
            "`n`r" + "Ticket Number:" + "`t`t" + $global:TicketTextBox.Text +
			"`n`r" + "Selected Solutions:" + "`t`t" + $global:SolutionsListBox.CheckedItems + 
            "`n`r" + "Selected Environment:" + "`t`t" + $global:EnvironmentComboBox.SelectedItem + 
            "`n`r" + "Backup Path:" + "`t`t" + $global:BackupPathTextBox.Text +
            "`n`r" + "Status Text:" + "`t`t" + $global:StatusText.Text + 
            "`n`r" + "Backed Up Sites:" + "`n`r" + "$($global:backedupsites|ft -autosize|out-string)" +
            "`n`r 1 = Solution Backed UP" + "`n`r 2 = Site and Solution Backed Up" + "`n`r 3 = Site, Solution and Web Config Backed Up"
        $global:DebugLabelText.refresh()
	}	
}

function updateBackupPathTextBox {
    if ($global:BackupPathTextBox.Text -eq "--[Backup Path]--"){} #do nothing 
    else{
        $global:BackupPathTextBox.Text="b:\$($global:landscape)\sharepoint\wfe01\$(get-date -f MM-dd-yyyy)\backups\$($global:TicketTextBox.Text)"#default backup path using global variables, also appears in $EnvironmentComboBox object
        $global:BackupPathTextBox.Refresh()
    }
}

#MainToolbar
$MainToolbar.Items.AddRange($fileToolStripMenuItem)
$MainToolbar.Location = new-object System.Drawing.Point(0, 0)
$MainToolbar.Name = "MainToolbar"
$MainToolbar.Size = new-object System.Drawing.Size(400, 24)
$MainToolbar.TabIndex = 0
$MainToolbar.Text = "menuStrip1"

#fileToolStripMenuItem
$fileToolStripMenuItem.DropDownItems.AddRange(@($resetToolStripMenuItem,$debugToolStripMenuItem,$exitToolStripMenuItem))
$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
$fileToolStripMenuItem.Size = new-object System.Drawing.Size(35, 20)
$fileToolStripMenuItem.Text = "&File"

#resetToolStripMenuItem
$resetToolStripMenuItem.Name = "resetToolStripMenuItem"
$resetToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$resetToolStripMenuItem.Text = "&Reset Form"
function OnClick_resetToolStripMenuItem{
    $SolutionsListBox.Items.Clear()
    $BackupPathTextBox.ResetText()
    $BackupPathTextBox.Text = "--[Backup Path]--"
    $StatusText.ResetText()
    $StatusText.Text = "Please Enter a Ticket # to get started."
    $TicketTextBox.ResetText()
    $TicketTextBox.Text = "Ticket #"
    $EnvironmentComboBox.ResetText()
    $EnvironmentComboBox.Text = "-Select Environment-"
	$backedupsites.clear()
	if ($debugmode -eq "1"){
		$Form.ClientSize = new-object System.Drawing.Size(400, 520)
        $MainToolbar.Size = new-object System.Drawing.Size(400, 24)
        $DebugLabelText.Dispose()
		$DebugMode=0
        Remove-Variable DebugMode
	}
    updateBackupPathTextBox
}
$resetToolStripMenuItem.Add_Click({OnClick_resetToolStripMenuItem})

#debugToolStripMenuItem
$debugToolStripMenuItem.Name = "debugToolStripMenuItem"
$debugToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$debugToolStripMenuItem.Text = "&Debug Mode"
$debugToolStripMenuItem.Add_Click({
    if (!($DebugMode)) {
        $Form.ClientSize = new-object System.Drawing.Size(800, 520)
        $MainToolbar.Size = new-object System.Drawing.Size(800, 24)
        $DebugLabelText = new-object System.Windows.Forms.Label
        $DebugLabelText.Size = new-object System.Drawing.Size(400, 500)
        $DebugLabelText.Location = new-object System.Drawing.Point(400,45)
        #$DebugLabelText.TextAlign = "MiddleCenter"
        $DebugLabelText.Forecolor = "Red"
        $DebugLabelText.Text = "Debug info will show up here:"
        $DebugMode=1
        $Form.Controls.AddRange(@($DebugLabelText))    
    }
    else {
        $Form.ClientSize = new-object System.Drawing.Size(400, 520)
        $MainToolbar.Size = new-object System.Drawing.Size(400, 24)
        $DebugLabelText.Dispose()
		$DebugMode=0
        Remove-Variable DebugMode
    }
})

#exitToolStripMenuItem
$exitToolStripMenuItem.Name = "exitToolStripMenuItem"
$exitToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$exitToolStripMenuItem.Text = "&Exit"
function OnClick_exitToolStripMenuItem {
$form.close()
}
$exitToolStripMenuItem.Add_Click({OnClick_exitToolStripMenuItem})

#SolutionsListBox
$SolutionsListBox = new-object System.Windows.Forms.CheckedListBox
$SolutionsListBox.Size = new-object System.Drawing.Size(400,400)
$SolutionsListBox.Location = new-object System.Drawing.Point(0,45)
$SolutionsListBox.CheckOnClick = $true
$SolutionsListBox.TabIndex = 3
$SolutionsListBox.add_click({
    if ($TicketTextBox.Text -eq "Ticket #"){$StatusText.Text = "Please Enter a Ticket Number First.";$StatusText.Refresh()}
    else {
        if ($SolutionsListBox.Items -ne ""){$StatusText.Text = "After selecting the solutions, select an environment.";$StatusText.Refresh()}
        else {$StatusText.Text = "Now click on the Load Solutions button.";$StatusText.Refresh()}    
    }
}) 
$SolutionsListBox.add_SelectedValueChanged({updateDebugLabels;$BackupPathTextBox.refresh()})

#BackupPathTextBox
$BackupPathTextBox = new-object System.Windows.Forms.TextBox
$BackupPathTextBox.Size = new-object System.Drawing.Size(400,20)
$BackupPathTextBox.Location = new-object System.Drawing.Point(0,460)
$BackupPathTextBox.TextAlign = "Center"
$BackupPathTextBox.Text = "--[Backup Path]--"

#TicketTextBox
$TicketTextBox = new-object System.Windows.Forms.TextBox
$TicketTextBox.Size = new-object System.Drawing.Size(60,20)
$TicketTextBox.Location = new-object System.Drawing.Point(0,25)
$TicketTextBox.TextAlign = "Center"
$TicketTextBox.Text = "Ticket #"
$TicketTextBox.TabIndex = 1
$TicketTextBox.add_click({
    if ($TicketTextBox.Text -eq "Ticket #") {$TicketTextBox.Clear()}
})
$TicketTextBox.add_leave({
    if ($TicketTextBox.Text -eq "") {$TicketTextBox.Text = "Ticket #"}
    if ($TicketTextBox.Text -eq "Ticket #"){$StatusText.Text = "Please Enter a Ticket Number First.";$StatusText.Refresh()}
    else {$StatusText.Text = "Now click on the Load Solutions button.";$StatusText.Refresh()}
    updateDebugLabels
    updateBackupPathTextBox
})

#StatusText
$StatusText = new-object System.Windows.Forms.Label
$StatusText.Size = new-object System.Drawing.Size(400, 20)
$StatusText.Location = new-object System.Drawing.Point(0,480)
$StatusText.Text = "Please Enter a Ticket # to get started."
$StatusText.TextAlign = "MiddleCenter"
$StatusText.Forecolor = "Red"

#ProgressBar
$ProgressBar = new-object System.Windows.Forms.ProgressBar
$ProgressBar.Size = new-object System.Drawing.Size(400, 20)
$ProgressBar.Location = new-object System.Drawing.Point(0,500)
$ProgressBar.Maximum = 100
$ProgressBar.Minimum = 0
$Progress = 0

#EnvironmentComboBox
$EnvironmentComboBox = new-object System.Windows.Forms.Combobox
$EnvironmentComboBox.size = new-object System.Drawing.Size(133, 20)
$EnvironmentComboBox.Location = new-object System.Drawing.Point(0, 440)
$EnvironmentComboBox.text = "-Select Environment-"
$EnvironmentComboBox.Items.AddRange(@("Local","DevInt","QA","Preview/DR","Production"))
$EnvironmentComboBox.add_SelectionChangeCommitted({
    if ($EnvironmentComboBox.SelectedItem -eq "Preview/DR") {$landscape = "DR"}
	else {
		if ($EnvironmentComboBox.SelectedItem -eq "Production") {$landscape = "Prod"}
		else {$landscape = $EnvironmentCombobox.SelectedItem}
	}	
	$BackupPathTextBox.Text="b:\$landscape\sharepoint\wfe01\$(get-date -f MM-dd-yyyy)\backups\$($TicketTextBox.Text)"#default backup path using global variables, also appears in updateDebugLabels function
	if ($TicketTextBox.Text -eq "Ticket #"){$StatusText.Text = "Please Enter a Ticket Number First.";$StatusText.Refresh()}
	else {
		if ($SolutionsListBox.Items -ne "") {$StatusText.Text = "Verify the backup path, and click the Backup Site(s) button";$StatusText.Refresh()}
		else {$StatusText.Text = "Please select the solutions you will be upgrading.";$StatusText.Refresh()} 
	}
    updateDebugLabels
})
$EnvironmentComboBox.TabIndex = 4

#LoadSolutionsButton
$LoadSolutionsButton = new-object Windows.Forms.Button
$LoadSolutionsButton.size = new-object System.Drawing.Size(340, 20)
$LoadSolutionsButton.Location = new-object System.Drawing.Point(60, 25)
$LoadSolutionsButton.TabIndex = 2
$LoadSolutionsButton.Text = "Load Solutions"
$LoadSolutionsButton.add_click(
{
    $SolutionsListBox.Items.Clear()
    $SolutionsListBox.BeginUpdate()
    foreach ($wsp in Get-SPSolution) {$SolutionsListBox.Items.Add($wsp.Name.ToString())}
    $SolutionsListBox.EndUpdate()
    if ($TicketTextBox.Text -eq "Ticket #"){$StatusText.Text = "Please Enter a Ticket Number First.";$StatusText.Refresh()}
    else {$StatusText.Text = "Please select the solutions you will be upgrading.";$StatusText.Refresh()}
})

#BackupSitesButton
$BackupSitesButton = new-object Windows.Forms.Button
$BackupSitesButton.Size = new-object System.Drawing.Size(134, 20)
$BackupSitesButton.Location = new-object System.Drawing.Point(133, 440)
$BackupSitesButton.Text = "Backup"
$BackupSitesButton.TabIndex = 5
$BackupSitesButton.add_click(
{
	$backedupsites.clear()
    if ($TicketTextBox.Text -eq "Ticket #") {$StatusText.Text = "Please Enter a Ticket Number First."}
    else {
        if (!($SolutionsListBox.CheckedItems)) {$StatusText.Text = "Please select the solutions you will be upgrading."}
        else {
            if (!($EnvironmentComboBox.SelectedItem)){$StatusText.Text = "Please select an environment."}
            else {
                $StatusText.Text = "Please be patient while we back up web applications that use this solution."
                $StatusText.Refresh()                    
                #make a hash table of sites that solutions are deployed to, in order to ensure we dont duplicate site backups
                $SolutionsListBox.CheckedItems|foreach {
                    $solution = Get-SPSolution $_.ToString()
                    $solution.DeployedWebApplications|Foreach {
                        $url = "$($_.URL)"
                        if ($backedupsites.contains($url)) {}#do nothing
                        else {$backedupsites += @{$url="0"};updateDebugLabels}#add it to hash table
                    }
                }
                #for each solution begin backups
                $SolutionsListBox.CheckedItems|foreach {
                    $solution = Get-SPSolution $_.ToString()
                    write-host $solution.Name
                    mkdir -force "$($BackupPathTextBox.Text)\Solutions"    
                    $path0 = "$($BackupPathTextBox.Text)\Solutions\$($solution.Name)"    
                    $solutionname = $solution.Name
					$jobname = Get-Random
					write-host "Now attempting the backup of $solutionname."
                    $StatusText.Text = "Now attempting the backup $solutionname." 
                    $StatusText.Refresh()  
                    start-job -Name $jobname -ArgumentList @($solutionname,$path0) -ScriptBlock {$solutionname=$args[0];$path0=$args[1];Add-PSSnapin "Microsoft.SharePoint.PowerShell";$solution = Get-SPSolution $solutionname;$solution.SolutionFile.SaveAs($path0)}
                    while ((get-job -name $jobname|select -expandproperty State) -eq "Running" -OR (get-job -name $jobname|select -expandproperty State) -eq "Stopping" -OR (get-job -name $jobname|select -expandproperty State) -eq "Suspending") {
                        $progress++
                        if ($progress -eq '100') {$progress = 0;$progress++}
                        $ProgressBar.Value = $progress
                        start-sleep 1
                    }
                    if ((get-job -Name $jobname|select -expandproperty State) -eq "Completed") {
                        $error.clear()
                        Receive-Job -Name $jobname -EA continue
                        $jobstatus=$error[0]
                        if (!($jobstatus)) {$solutionbackup="1";$ProgressBar.Value="100";write-host "Backup of $solutionname completed successfully.";$StatusText.Text = "Backup of $solutionname completed successfully.";$StatusText.Refresh()}
                        else {$solutionbackup="0";[void][System.Windows.Forms.MessageBox]::Show($solutionname + ":`t" + $jobstatus)}    
                    }
                    start-sleep 1
                    $Progress= 0
                    if (!($solution.DeployedWebApplications)){write-host "$solution.Name is not deployed to any web application";$StatusText.Text = "$solution.Name is not deployed to any web application";$StatusText.Refresh()}
                    else {
                        $solution.DeployedWebApplications|Foreach {
							$url = "$($_.URL)"
                            if ($solutionbackup = 0) {$backedupsites.set_item($url,"Solution Backup Failed");updateDebugLabels}
                            else {
								$backedupsites.set_item($url,"1")
                                updateDebugLabels								
								if ($backedupsites.get_item($url) -eq "1"){
									write-host "Starting site content backup of $url"
									mkdir -force "$($BackupPathTextBox.Text)\SiteContent"
									$path1 ="$($BackupPathTextBox.Text)\SiteContent\$url.bak" -replace "http://" -replace "/"
									$jobname = Get-Random
									$StatusText.Text = "Now attempting the site content backup of $url."
									$StatusText.Refresh()  
									start-job -Name $jobname -ArgumentList @($url,$path1) -ScriptBlock {$url=$args[0];$path1=$args[1];Add-PSSnapin "Microsoft.SharePoint.PowerShell";Backup-SPSite -identity "$url" -path "$path1" -confirm:$false -force} 
									while ((get-job -name $jobname|select -expandproperty State) -eq "Running" -OR (get-job -name $jobname|select -expandproperty State) -eq "Stopping" -OR (get-job -name $jobname|select -expandproperty State) -eq "Suspending") {
										$progress++
										if ($progress -eq '100') {$progress = 0;$progress++}
										$ProgressBar.Value = $progress
										start-sleep 1
									}
									if ((get-job -Name $jobname|select -expandproperty State) -eq "Completed") {
										$error.clear()
										Receive-Job -Name $jobname -EA continue
										$jobstatus=$error[0]
										if (!($jobstatus)) {$backedupsites.set_item($url,"2");$ProgressBar.Value=100;write-host "Backup of $url completed successfully.";$StatusText.Text = "Backup of $url completed successfully.";$StatusText.Refresh();updateDebugLabels}
										else {$backedupsites.set_item($url,"Error Backing Up Site");updateDebugLabels;[void][System.Windows.Forms.MessageBox]::Show($url + ":`t" + $jobstatus)}    
									}
									if ($backedupsites.get_item($url) -eq "2"){
										write-host "Starting config file backup of $url"
										$path2 = "$($($BackupPathTextBox.Text) + "\WebConfigs")"
										mkdir $path2 -force
                                        foreach ($webconfigfullpath in (gci -recurse (((Get-SPWebApplication $url).IisSettings.Values)|select Path).Path.Fullname |where-object {$_.Name -ilike '*.config' -and $_.FullName -notlike '*wpresources*'}|select-object -expandproperty FullName)){
												mkdir $path2\$(($webconfigfullpath).replace("$(((((Get-SPWebApplication $url).IisSettings.Values)|select Path).Path.Fullname).substring(0, (((((Get-SPWebApplication $url).IisSettings.Values)|select Path).Path.Fullname).lastindexOf('\'))))\","") -replace "\\\w{0,99}\.config","") -force
												cp $webconfigfullpath $path2\$(($webconfigfullpath).replace("$(((((Get-SPWebApplication $url).IisSettings.Values)|select Path).Path.Fullname).substring(0, (((((Get-SPWebApplication $url).IisSettings.Values)|select Path).Path.Fullname).lastindexOf('\'))))\","") -replace "\\\w{0,99}\.config","") -force
										}
										$backedupsites.set_item($url,"3")
										updateDebugLabels		
									}
								}
								else {write-host "Skipping site content backup of $url because it has already been completed"}#do nothing because the site is already backed up
								start-sleep 1
								$Progress=0
								$ProgressBar.Value=0
								updateDebugLabels
							}
						}	
                    }
					updateDebugLabels
				}
                $StatusText.Text = "Backups have completed."
                $StatusText.Refresh()
				write-host "Backups have completed."
				updateDebugLabels
			}
			updateDebugLabels
		}
		updateDebugLabels
	}
updateDebugLabels
})

#ShowBackupsButton
$ShowBackupsButton = new-object Windows.Forms.Button
$ShowBackupsButton.size = new-object System.Drawing.Size(133, 20)
$ShowBackupsButton.Location = new-object System.Drawing.Point(267, 440)
$ShowBackupsButton.TabIndex = 2
$ShowBackupsButton.Text = "Show Backups"
$ShowBackupsButton.TabIndex = 6
$ShowBackupsButton.add_click(
{
    if (test-path $($BackupPathTextBox.Text)){start-process explorer "$($BackupPathTextBox.Text)"}
    else {
        $StatusText.Text = "Please verify the backup path and try again."
        $StatusText.Refresh()
		write-host "Please verify the backup path and try again."
		updateDebugLabels
    }
})


$Form.Controls.AddRange(@($ShowBackupsButton,$ProgressBar,$MainToolbar,$StatusText,$TicketTextBox,$EnvironmentComboBox,$LoadSolutionsButton,$BackupSitesButton,$SolutionsListBox,$BackupPathTextBox))
$Form.MainMenuStrip = $MailToolbar
$Form.Name = "MenuForm"
$Form.Text = "SharePoint 2010 Sites & Solutions Backup Application"
function OnFormClosing_MenuForm($Sender,$e){ 
    # $this represent sender (object)
    # $_ represent  e (eventarg)

    # Allow closing
    ($_).Cancel= $False
}
$Form.Add_FormClosing( { OnFormClosing_MenuForm $Form $EventArgs} )
$Form.Add_Shown({$Form.Activate()})

$Form.ShowDialog()
#Free ressources
$Form.Dispose()
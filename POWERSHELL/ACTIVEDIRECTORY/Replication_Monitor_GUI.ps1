<#
	.SYNOPSIS
		Active Directory Replication Monitor GUI  ( replacement for  ADREPLSTATUS from MS )

	.DESCRIPTION
		Script based on parsing of repadmin /showrepl output
		
	.EXAMPLE
		PS C:\> .\Replication_Monitor_GUI.ps1

	.OUTPUTS
		Can save results from view to csv file

	.NOTES
		Run With Powershell 4 and Domain Admin Rights. 
		Tested with AD2012 ( Windows Server 2012 R2 )

	.NOTES 
        Take It, Hold It, Love It 
		
	.NOTES 
    	Please contact me if any issues with script. tnx
 
    .LINK 
        Author : Andrew V. Golubenkoff (andrew.golubenkoff@outlook.com) 

	.VER
		v1.5 Initial
                v1.6 Fixed Powershell version Check
#>




#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(1008, 634)
$Form1.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$Form1.MaximizeBox = $false
$Form1.MinimizeBox = $false
$Form1.ShowIcon = $false
$Form1.Text = "Active Directory Replication Status"
#~~< Button_Refresh >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button_Refresh = New-Object System.Windows.Forms.Button
$Button_Refresh.AutoSize = $true
$Button_Refresh.Location = New-Object System.Drawing.Point(638, 25)
$Button_Refresh.Size = New-Object System.Drawing.Size(143, 23)
$Button_Refresh.TabIndex = 2
$Button_Refresh.Text = "Refresh Replication Status"
$Button_Refresh.UseVisualStyleBackColor = $true
$Button_Refresh.add_Click({Button_RefreshClick($Button_Refresh)})
$Form1.AcceptButton = $Button_Refresh
#~~< Label4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label4 = New-Object System.Windows.Forms.Label
$Label4.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$Label4.Location = New-Object System.Drawing.Point(13, 599)
$Label4.Size = New-Object System.Drawing.Size(983, 23)
$Label4.TabIndex = 8
$Label4.Text = "Created by Andrew V. Golubenkoff"
#~~< ProgressBar1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar1.Location = New-Object System.Drawing.Point(12, 89)
$ProgressBar1.Size = New-Object System.Drawing.Size(984, 23)
$ProgressBar1.TabIndex = 7
$ProgressBar1.Text = ""
#~~< CheckBox_erroronly >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CheckBox_erroronly = New-Object System.Windows.Forms.CheckBox
$CheckBox_erroronly.Location = New-Object System.Drawing.Point(638, 51)
$CheckBox_erroronly.Size = New-Object System.Drawing.Size(104, 24)
$CheckBox_erroronly.TabIndex = 6
$CheckBox_erroronly.Text = "Errors Only"
$CheckBox_erroronly.UseVisualStyleBackColor = $true
$CheckBox_erroronly.ForeColor = [System.Drawing.Color]::Firebrick
#~~< DataGridView1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DataGridView1 = New-Object System.Windows.Forms.DataGridView
$DataGridView1.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$DataGridView1.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DataGridView1.Location = New-Object System.Drawing.Point(12, 120)
$DataGridView1.Size = New-Object System.Drawing.Size(984, 470)
$DataGridView1.TabIndex = 5
$DataGridView1.Text = ""
$DataGridView1.BackgroundColor = [System.Drawing.Color]::AntiqueWhite
#~~< Label3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Location = New-Object System.Drawing.Point(410, 51)
$Label3.Size = New-Object System.Drawing.Size(193, 23)
$Label3.TabIndex = 4
$Label3.Text = "Filter"
$Label3.ForeColor = [System.Drawing.SystemColors]::HotTrack
$Label3.add_Click({Label2Click($Label3)})
#~~< Label2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Point(211, 51)
$Label2.Size = New-Object System.Drawing.Size(193, 23)
$Label2.TabIndex = 4
$Label2.Text = "Filter by Target"
$Label2.ForeColor = [System.Drawing.SystemColors]::HotTrack
$Label2.add_Click({Label2Click($Label2)})
#~~< Label1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Location = New-Object System.Drawing.Point(12, 51)
$Label1.Size = New-Object System.Drawing.Size(193, 23)
$Label1.TabIndex = 4
$Label1.Text = "Filter by Source"
$Label1.ForeColor = [System.Drawing.SystemColors]::HotTrack
#~~< TextBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Point(410, 27)
$TextBox1.Size = New-Object System.Drawing.Size(222, 20)
$TextBox1.TabIndex = 3
$TextBox1.Text = ""
$TextBox1.add_KeyPress({TextBox1KeyPress($TextBox1)})
$TextBox1.add_TextChanged({TextBox1TextChanged($TextBox1)})
#~~< Button_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button_Export = New-Object System.Windows.Forms.Button
$Button_Export.AutoSize = $true
$Button_Export.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$Button_Export.Location = New-Object System.Drawing.Point(838, 25)
$Button_Export.Size = New-Object System.Drawing.Size(49, 23)
$Button_Export.TabIndex = 2
$Button_Export.Text = "Export"
$Button_Export.UseVisualStyleBackColor = $true
$Button_Export.ForeColor = [System.Drawing.SystemColors]::ActiveCaption
$Button_Export.add_Click({Button_ExportClick($Button_Export)})
#~~< Button_Clear >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Button_Clear = New-Object System.Windows.Forms.Button
$Button_Clear.AutoSize = $true
$Button_Clear.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](204)))
$Button_Clear.Location = New-Object System.Drawing.Point(787, 25)
$Button_Clear.Size = New-Object System.Drawing.Size(45, 23)
$Button_Clear.TabIndex = 2
$Button_Clear.Text = "Clear"
$Button_Clear.UseVisualStyleBackColor = $true
$Button_Clear.ForeColor = [System.Drawing.SystemColors]::ActiveCaption
$Button_Clear.add_Click({Button_ClearClick($Button_Clear)})
#~~< ComboBox2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox2 = New-Object System.Windows.Forms.ComboBox
$ComboBox2.FormattingEnabled = $true
$ComboBox2.Location = New-Object System.Drawing.Point(211, 27)
$ComboBox2.SelectedIndex = -1
$ComboBox2.Size = New-Object System.Drawing.Size(193, 21)
$ComboBox2.TabIndex = 1
$ComboBox2.Text = ""
$ComboBox2.add_KeyPress({ComboBox2KeyPress($ComboBox2)})
$ComboBox2.add_SelectedValueChanged({ComboBox2SelectedValueChanged($ComboBox2)})
#~~< ComboBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ComboBox1 = New-Object System.Windows.Forms.ComboBox
$ComboBox1.FormattingEnabled = $true
$ComboBox1.Location = New-Object System.Drawing.Point(12, 27)
$ComboBox1.SelectedIndex = -1
$ComboBox1.Size = New-Object System.Drawing.Size(193, 21)
$ComboBox1.TabIndex = 1
$ComboBox1.Text = ""
$ComboBox1.add_KeyPress({ComboBox1KeyPress($ComboBox1)})
$ComboBox1.add_SelectedValueChanged({ComboBox1SelectedValueChanged($ComboBox1)})
#~~< MainMenu1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu1 = New-Object System.Windows.Forms.MenuStrip
$MainMenu1.Location = New-Object System.Drawing.Point(0, 0)
$MainMenu1.Size = New-Object System.Drawing.Size(1008, 24)
$MainMenu1.TabIndex = 0
$MainMenu1.Text = "MainMenu1"
#~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$FileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
$FileToolStripMenuItem.Text = "File"
#~~< RefreshToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RefreshToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$RefreshToolStripMenuItem.Size = New-Object System.Drawing.Size(172, 22)
$RefreshToolStripMenuItem.Text = "Refresh Forest Info"
$RefreshToolStripMenuItem.add_Click({RefreshToolStripMenuItemClick($RefreshToolStripMenuItem)})
#~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ExitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExitToolStripMenuItem.Size = New-Object System.Drawing.Size(172, 22)
$ExitToolStripMenuItem.Text = "Exit"
$ExitToolStripMenuItem.add_Click({ExitToolStripMenuItemClick($ExitToolStripMenuItem)})
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($RefreshToolStripMenuItem, $ExitToolStripMenuItem)))
#~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutToolStripMenuItem.Size = New-Object System.Drawing.Size(52, 20)
$AboutToolStripMenuItem.Text = "About"
$AboutToolStripMenuItem.add_Click({AboutToolStripMenuItemClick($AboutToolStripMenuItem)})
$MainMenu1.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@($FileToolStripMenuItem, $AboutToolStripMenuItem)))
$Form1.Controls.Add($Label4)
$Form1.Controls.Add($ProgressBar1)
$Form1.Controls.Add($CheckBox_erroronly)
$Form1.Controls.Add($DataGridView1)
$Form1.Controls.Add($Label3)
$Form1.Controls.Add($Label2)
$Form1.Controls.Add($Label1)
$Form1.Controls.Add($TextBox1)
$Form1.Controls.Add($Button_Export)
$Form1.Controls.Add($Button_Clear)
$Form1.Controls.Add($Button_Refresh)
$Form1.Controls.Add($ComboBox2)
$Form1.Controls.Add($ComboBox1)
$Form1.Controls.Add($MainMenu1)
#~~< SaveFileDialog1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SaveFileDialog1 = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog1.DefaultExt = "html"
$SaveFileDialog1.FileName = "ReplicationStatus"
$SaveFileDialog1.Filter = ""+[char]34+"CSV Files|*.csv|All Files|*.*"+[char]34
$SaveFileDialog1.ShowHelp = $true

#endregion

#region Custom Code

#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$Form1.ShowDialog()
}

#endregion

#endregion

#region Event Handlers

	
	# My Functions
	function AG_ForestInfoLoad
	{
	$Global:myForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest() 
	$Global:dclist = $Global:myForest.Sites | % { $_.Servers } 
 	$ComboBox1.Items.Clear()
	$ComboBox2.Items.Clear()
	$ComboBox1.Items.AddRange($(($Global:dclist ).Name))
	$ComboBox2.Items.AddRange($($(($Global:dclist ).Name) -replace "\..*") )
	}
	
	function AG_ReplicationStatus
	{
	param([string]$filter,[string]$direction)

	if ($Global:FilterText){$filter = $Global:FilterText }
	elseif ($ComboBox1.SelectedValue.lenght -gt 1){$filter = $ComboBox1.SelectedValue }
	
	
	$date = Get-Date -Format "dd.MM.yyyy_HH.mm"
	$global:array = @() 
	$c=0
	$ProgressBar1.Visible = $true
		foreach ($dcname in $dclist){ 
		    $c+=1
			$ProgressBar1.Value=$($c*100/$dclist.Count)
			$source_dc_fqdn = ($dcname.Name).tolower() 
		    $ad_partition_list = repadmin /showrepl $source_dc_fqdn | Select-String "dc=" 
			
			
		    foreach ($ad_partition in $ad_partition_list) { 
			
			
		        [Array]$NewArray=$NULL 
		        $result = repadmin /showrepl $source_dc_fqdn $ad_partition 
		        $result = $result | where { ([string]::IsNullOrEmpty(($result[$_]))) } 
		        $index_array_dst = 0..($result.Count - 1) | where { $result[$_] -like "*via RPC" } 
		        foreach ($index in $index_array_dst){ 
		            $dst_dc = ($result[$index]).trim() 
		            $next_index = [array]::IndexOf($index_array_dst,$index) + 1 
		            $next_index_msg = $index_array_dst[$next_index] 
		            $msg = "" 
		            if ($index -lt $index_array_dst[-1]){ 
		                $last_index = $index_array_dst[$next_index] 
		            } 
		            else { 
		                $last_index = $result.Count 
		            } 
		            
		            for ($i=$index+1;$i -lt $last_index; $i++){ 
		                if (($msg -eq "") -and ($result[$i])) { 
		                    $msg += ($result[$i]).trim() 
		                } 
		                else { 
		                    $msg += " / " + ($result[$i]).trim() 
		                } 
		            } 
				
					if  ( $($($msg -split "\/")[3]) -match "access was denied"){$status = "Access Denied"}
					elseif ($($($msg -split "\/")[1]) -match "successful"){$status = "OK"}
					else{$status = "ERROR"}
					$Properties = @{source_dc=$source_dc_fqdn;NC=$ad_partition;destination_site=$($dst_dc -replace "\\.*");destination_dc=$($dst_dc -replace ".*\\" -replace "\s.*") ;status=$status;time=$(($msg -split "\s")[8] + " " + ($msg -split "\s")[9]);repl_status=$msg} 
		            $Newobject = New-Object PSObject -Property $Properties 
		            $global:array +=$newobject 
		        } 
		    } 
		} 
		if ($CheckBox_erroronly.Checked -eq $true){
			$global:array = $($global:array |  where {$_.repl_status -notlike "*successful*"})
		}
		
		if ($filter)
		{
			switch ($Global:Direction)
			{
			"Source" { $global:array = $($global:array |  where {$_.source_dc -match $filter }) }
			"Target" { $global:array = $($global:array |  where {$_.destination_dc -match $filter }) }
			"Filter"  {$global:array = $($global:array |  where {$_ -match $filter }) }
			}
			
		
		}
		
	} # END ReplicationStatus Function
	
	function AG_FillForm
	{
	$list = New-Object System.collections.ArrayList
	if ($array -eq $null){
	[System.Windows.Forms.MessageBox]::Show("No Results Found") 
	}
	elseif ($array -eq $null -and $CheckBox_erroronly.Checked )
	{
	[System.Windows.Forms.MessageBox]::Show("No Results with Errors") 
	}
	else
	{
		$list.AddRange( $($array | select source_dc,destination_dc,destination_site,NC,status,time,repl_status) )
		$dataGridView1.ColumnHeadersVisible = $true
		$dataGridView1.AutoSizeColumnsMode = "AllCells"
		$dataGridView1.DataSource = $list

			foreach ($Row in $DataGridView1.Rows)
			{  
				if ($Row.DataBoundItem.Status -eq "OK")
				{
					$dataGridView1.Rows[$row.Index].Cells[4].Style.ForeColor = "Green";
				}
				elseif ($Row.DataBoundItem.Status -eq "Access Denied")
				{
					$dataGridView1.Rows[$row.Index].Cells[4].Style.ForeColor = "DarkOrange";
				}
				else
				{
					$dataGridView1.Rows[$row.Index].Cells[4].Style.ForeColor = "Red";
				}
				
			}
		}
	} # END AG_FillFormFunction

	function AG_ClearForm
	{
	$array = @()
	$list = New-Object System.collections.ArrayList
	$list.AddRange($array)
	$dataGridView1.ColumnHeadersVisible = $true
	$dataGridView1.DataSource = $list
	$ProgressBar1.Value = 0
	$ProgressBar1.Visible = $false
	$ComboBox1.Items.Clear()
	$ComboBox2.Items.Clear()
	$ComboBox1.SelectedItem = $null
	$ComboBox2.SelectedItem = $null
	$ComboBox1.Text = $null
	$ComboBox2.Text = $null
	$TextBox1.Text = $null
	$Global:FilterText =  $null
	$Global:Direction = $null
	$Global:REFRESH = $true
	}

function Button_RefreshClick( $object ){
	#AG_ClearForm
	AG_ReplicationStatus -filter $Global:Direction -direction $Global:Direction
	AG_FillForm

$Global:REFRESH = $true	
}

function RefreshToolStripMenuItemClick( $object ){
AG_ClearForm
AG_ForestInfoLoad
}

function AboutToolStripMenuItemClick( $object ){
$About = "

Created by Andrew V. Golubenkoff
[ Andrew.Golubenkoff@outlook.com ]

Based on repadmin /showrepl output parsing

Start util with Domain Admin or equal rights.

P.S 
If you find any problems please let me know.
Also if you need some additional functionality - mail me :)


"
[System.Windows.Forms.MessageBox]::Show($About) 

}

function ExitToolStripMenuItemClick( $object ){
$Form1.Dispose()
}


function TextBox1KeyPress( $object ){
    if ($object.KeyChar -eq 13) { $Button_Refresh.PerformClick() }
}

function TextBox1TextChanged( $object ){
$Global:Direction = "Filter"
[bool]$Global:REFRESH = $false
$Global:FilterText = $TextBox1.Text
if ($Global:FilterText.length -eq 0)
{
$Global:FilterText = $null
$Global:Direction = $null
}

$ComboBox1.SelectedItem = $null
$ComboBox2.SelectedItem = $null
$ComboBox1.Text = $null
$ComboBox2.Text = $null
}

function ComboBox2SelectedValueChanged( $object ){
$Global:Direction = "Target"
$Global:FilterText = $ComboBox2.Text
if ($Global:REFRESH)
	{
		AG_ReplicationStatus -filter $ComboBox2.Text
		AG_FillForm
		$TextBox1.Text =  $null
	}
}

function ComboBox1SelectedValueChanged( $object ){

$Global:Direction = "Source"
$Global:FilterText = $ComboBox1.Text


	if ($Global:REFRESH)
	{
		AG_ReplicationStatus -filter $ComboBox1.Text
		AG_FillForm
		$TextBox1.Text =  $null
	}
}

function Button_ClearClick( $object ){
AG_ClearForm
AG_ForestInfoLoad
}
Button_ClearClick

function export-DGV2CSV ([Windows.Forms.DataGridView] $grid, [String] $File)
{
  if ($grid.RowCount -eq 0) { return } # nothing to do
  
  $row = New-Object Windows.Forms.DataGridViewRow
  $sw  = new-object System.IO.StreamWriter($File)
      
  # write header line
  $sw.WriteLine( ($grid.Columns | % { $_.HeaderText } ) -join ';' )

  # dump values
  $grid.Rows | % {
    $sw.WriteLine(
      ($_.Cells | % { $_.Value }) -join ';'
      )
    }
  $sw.Close()
}


function Button_ExportClick( $object )
{
	
$SaveFileDialog1.FileName = (Get-Date -Format "dd.MM.yyyy_HH.mm_")+ $Global:myForest.Name + "_" + $SaveFileDialog1.FileName
$SaveFileDialog1.ShowDialog()

$Filename = $SaveFileDialog1.FileName
$SaveFileDialog1

export-DGV2CSV -grid $DataGridView1 -File $Filename
}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
        return $principal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    } catch {
        throw "Failed to determine if the current user has elevated privileges. The error was: '{0}'." -f $_
    }
}

$isPS4 = $PSVersionTable.PSVersion.Major
if ($isPS4 -lt 4)
{
$message = "Run Script only with Powershell v.4 and higher! Current Version: $isPS4 "
[System.Windows.Forms.MessageBox]::Show($message,"!!! Warning !!!") 
break
}

$isAdmin = Test-IsAdmin
if ($isAdmin -eq $false)
{
$message = "Run Script with elevated privileges and Domain Admin Rights"
[System.Windows.Forms.MessageBox]::Show($message,"!!! Warning !!!") 
}


Main # This call must remain below all other event functions

#endregion


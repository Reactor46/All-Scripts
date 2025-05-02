#----------------------------------------------
#region Import Assemblies
#----------------------------------------------
[void][Reflection.Assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][Reflection.Assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
#endregion Import Assemblies

#Define a Param block to use custom parameters in the project
#Param ($CustomParameter)

function Main {
	Param ([String]$Commandline)
	#Note: This function starts the application
	#Note: $Commandline contains the complete argument string passed to the packager 
	#Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
	#Note: To get the console output in the Packager (Forms Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
	#TODO: Initialize and add Function calls to forms
	
	if((Call-MainForm_pff) -eq "OK")
	{
		
	}
	
	$global:ExitCode = 0 #Set the exit code for the Packager
}
#endregion Source: Startup.pfs

#region Source: MainForm.pff
function Call-MainForm_pff
{
	#----------------------------------------------
	#region Import the Assemblies
	#----------------------------------------------
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#endregion Import Assemblies

	#----------------------------------------------
	#region Generated Form Objects
	#----------------------------------------------
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$MainForm = New-Object 'System.Windows.Forms.Form'
	$buttonExecute = New-Object 'System.Windows.Forms.Button'
	$datagridviewMailboxes = New-Object 'System.Windows.Forms.DataGridView'
	$buttonCancel = New-Object 'System.Windows.Forms.Button'
	$tabcontrol1 = New-Object 'System.Windows.Forms.TabControl'
	$tabScope = New-Object 'System.Windows.Forms.TabPage'
	$grpCredentials = New-Object 'System.Windows.Forms.GroupBox'
	$txtServer = New-Object 'System.Windows.Forms.TextBox'
	$labelServer = New-Object 'System.Windows.Forms.Label'
	$buttonConnect = New-Object 'System.Windows.Forms.Button'
	$chkCurrentUser = New-Object 'System.Windows.Forms.CheckBox'
	$labelCurrentUser = New-Object 'System.Windows.Forms.Label'
	$txtUser = New-Object 'System.Windows.Forms.TextBox'
	$txtPassword = New-Object 'System.Windows.Forms.MaskedTextBox'
	$labelUser = New-Object 'System.Windows.Forms.Label'
	$labelPassword = New-Object 'System.Windows.Forms.Label'
	$grpScope = New-Object 'System.Windows.Forms.GroupBox'
	$btnLoadMailboxes = New-Object 'System.Windows.Forms.Button'
	$labelMailboxes = New-Object 'System.Windows.Forms.Label'
	$comboDatabase = New-Object 'System.Windows.Forms.ComboBox'
	$labelDatabase = New-Object 'System.Windows.Forms.Label'
	$comboServer = New-Object 'System.Windows.Forms.ComboBox'
	$comboDAG = New-Object 'System.Windows.Forms.ComboBox'
	$chkEnterprise = New-Object 'System.Windows.Forms.CheckBox'
	$labelCluster = New-Object 'System.Windows.Forms.Label'
	$labelEnterprise = New-Object 'System.Windows.Forms.Label'
	$labelDAG = New-Object 'System.Windows.Forms.Label'
	$tabAbout = New-Object 'System.Windows.Forms.TabPage'
	$richtextAbout = New-Object 'System.Windows.Forms.RichTextBox'
	$timerFadeIn = New-Object 'System.Windows.Forms.Timer'
	$tooltipAll = New-Object 'System.Windows.Forms.ToolTip'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	function PopulateDAGMailboxServers
	{  
	    Load-ComboBox $comboServer ''
	    Load-ComboBox $comboDatabase ''
	    $Script:MailboxServerScope = @()
	    foreach ($Server in @(Get-DatabaseAvailabilityGroup $comboDAG.Text|foreach {$_.Servers}))
	    {
	        Load-ComboBox $comboServer $Server.Name -Append
	        $Script:MailboxServerScope += $Server.Name
		}
	}
	
	function PopulateDAGs
	{  
	    Load-ComboBox $comboDatabase ''
	    Load-ComboBox $comboServer ''
	    foreach ($DAG in (Get-DatabaseAvailabilityGroup))
	    {
	        Load-ComboBox $comboDAG $DAG.Name -Append
		}
	}
	
	function PopulateMailboxServers
	{
	    Load-ComboBox $comboServer ''
	    Load-ComboBox $comboDatabase ''
	    $Script:MailboxServerScope = @()
	    foreach ($Server in ((Get-ExchangeServer | Where {$_.ServerRole -match 'Mailbox'})))
	    {
	        Load-ComboBox $comboServer $Server.Name -Append
	        $Script:MailboxServerScope += $Server.Name
		}
	}
	
	function PopulateDatabases
	{ 
	    Load-ComboBox $comboDatabase ''
	    $Script:MailboxDatabaseScope = @()
	    foreach ($Database in (Get-MailboxDatabase -Server $comboServer.Text))
	    {
	        Load-ComboBox $comboDatabase $Database.Name -Append
	        $Script:MailboxDatabaseScope += $Database.Name
		}
		
	}
	
	function PopulateMailboxes
	{
	    If ($chkEnterprise)
	    {
		}
	    elseif ($datagridviewMailboxes.IsSelected)
	    {
		}
	    elseif ($comboDatabase.Text -ne '')
	    {
		}
	    
	}
	
	$form1_FadeInLoad={
	
	
	    #Start the Timer to Fade In
		$timerFadeIn.Start()
		$MainForm.Opacity = 0
	
	    If (SnapinsAvailable)
	    {
	        If (LoadSnapins)
	        {
	            $MainForm.Text = "Exchange Wizard (Online)"
			}
	        else
	        {
	            $MainForm.Text = "Exchange Wizard (Offline - Snapin available)"
			}
		}
	    else
	    {
	        $MainForm.Text = "Exchange Wizard (Offline - Snapin not available)"
		}
	}
	
	$timerFadeIn_Tick={
		#Can you see me now?
		if($MainForm.Opacity -lt 1)
		{
			$MainForm.Opacity += 0.1
			
			if($MainForm.Opacity -ge 1)
			{
				#Stop the timer once we are 100% visible
				$timerFadeIn.Stop()
			}
		}
	}
	
	#region Control Helper Functions
	function Load-DataGridView
	{
		<#
		.SYNOPSIS
			This functions helps you load items into a DataGridView.
	
		.DESCRIPTION
			Use this function to dynamically load items into the DataGridView control.
	
		.PARAMETER  DataGridView
			The ComboBox control you want to add items to.
	
		.PARAMETER  Item
			The object or objects you wish to load into the ComboBox's items collection.
		
		.PARAMETER  DataMember
			Sets the name of the list or table in the data source for which the DataGridView is displaying data.
	
		#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.DataGridView]$DataGridView,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Item,
		    [Parameter(Mandatory=$false)]
			[string]$DataMember
		)
		$DataGridView.SuspendLayout()
		$DataGridView.DataMember = $DataMember
		
		if ($Item -is [System.ComponentModel.IListSource]`
		-or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView] )
		{
			$DataGridView.DataSource = $Item
		}
		else
		{
			$array = New-Object System.Collections.ArrayList
			
			if ($Item -is [System.Collections.IList])
			{
				$array.AddRange($Item)
			}
			else
			{	
				$array.Add($Item)	
			}
			$DataGridView.DataSource = $array
		}
		
		$DataGridView.ResumeLayout()
	}
	
	function Load-ComboBox 
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ComboBox.
	
		.DESCRIPTION
			Use this function to dynamically load items into the ComboBox control.
	
		.PARAMETER  ComboBox
			The ComboBox control you want to add items to.
	
		.PARAMETER  Items
			The object or objects you wish to load into the ComboBox's Items collection.
	
		.PARAMETER  DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER  Append
			Adds the item(s) to the ComboBox without clearing the Items collection.
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Load-ComboBox $combobox1 "Red" -Append
			Load-ComboBox $combobox1 "White" -Append
			Load-ComboBox $combobox1 "Blue" -Append
		
		.EXAMPLE
			Load-ComboBox $combobox1 (Get-Process) "ProcessName"
	#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.ComboBox]$ComboBox,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Items,
		    [Parameter(Mandatory=$false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if(-not $Append)
		{
			$ComboBox.Items.Clear()	
		}
		
		if($Items -is [Object[]])
		{
			$ComboBox.Items.AddRange($Items)
		}
		elseif ($Items -is [Array])
		{
			$ComboBox.BeginUpdate()
			foreach($obj in $Items)
			{
				$ComboBox.Items.Add($obj)	
			}
			$ComboBox.EndUpdate()
		}
		else
		{
			$ComboBox.Items.Add($Items)	
		}
	
		$ComboBox.DisplayMember = $DisplayMember	
	}
	
	function Load-ListBox 
	{
	<#
		.SYNOPSIS
			This functions helps you load items into a ListBox or CheckedListBox.
	
		.DESCRIPTION
			Use this function to dynamically load items into the ListBox control.
	
		.PARAMETER  ListBox
			The ListBox control you want to add items to.
	
		.PARAMETER  Items
			The object or objects you wish to load into the ListBox's Items collection.
	
		.PARAMETER  DisplayMember
			Indicates the property to display for the items in this control.
		
		.PARAMETER  Append
			Adds the item(s) to the ListBox without clearing the Items collection.
		
		.EXAMPLE
			Load-ListBox $ListBox1 "Red", "White", "Blue"
		
		.EXAMPLE
			Load-ListBox $listBox1 "Red" -Append
			Load-ListBox $listBox1 "White" -Append
			Load-ListBox $listBox1 "Blue" -Append
		
		.EXAMPLE
			Load-ListBox $listBox1 (Get-Process) "ProcessName"
	#>
		Param (
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			[System.Windows.Forms.ListBox]$listBox,
			[ValidateNotNull()]
			[Parameter(Mandatory=$true)]
			$Items,
		    [Parameter(Mandatory=$false)]
			[string]$DisplayMember,
			[switch]$Append
		)
		
		if(-not $Append)
		{
			$listBox.Items.Clear()	
		}
		
		if($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
		{
			$listBox.Items.AddRange($Items)
		}
		elseif ($Items -is [Array])
		{
			$listBox.BeginUpdate()
			foreach($obj in $Items)
			{
				$listBox.Items.Add($obj)
			}
			$listBox.EndUpdate()
		}
		else
		{
			$listBox.Items.Add($Items)	
		}
	
		$listBox.DisplayMember = $DisplayMember	
	}
	#endregion
	
	$buttonSave_Click={
	    Set-SaveData
	    if (Save-Config)
	    {
	        Set-ConnectionInfoLabel
	        #[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("Configuration saved.","All went well!")
		}
	}
	
	$buttonRun_Click={
		Set-SaveData
	    if (Save-Config)
	    {
	    	# Run the script
	    	$script = $ScriptDirectory + $StarterScript
	    	&$script
		}
	}
	
	$buttonBrowse_Click2={
	
		if($openfiledialog1.ShowDialog() -eq 'OK')
		{
			$txtReportFolder.Text = $openfiledialog1.FileName
		}
	}
	
	$chkSaveLocally_CheckedChanged={
		if ($chkSaveLocally.Checked)
		{
			$buttonBrowseFolder.Enabled = $true
	        $txtReportFolder.Enabled = $true
			$txtReportName.Enabled = $true
		}
		else
		{
			$buttonBrowseFolder.Enabled = $false
	        $txtReportFolder.Enabled = $false
			$txtReportName.Enabled = $false
		}
	}
	
	$buttonBrowseFolder_Click={
		if($folderbrowserdialog1.ShowDialog() -eq 'OK')
		{
			$txtReportFolder.Text = $folderbrowserdialog1.SelectedPath
		}
	}
	
	$chkEmailReport_CheckedChanged={
		if ($chkEmailReport.Checked)
		{
	        $txtEmailSubject.Enabled = $true
	        $txtEmailRecipient.Enabled = $true
	        $txtEmailSender.Enabled = $true
	        $txtSMTPServer.Enabled = $true
		}
		else
		{
	        $txtEmailSubject.Enabled = $false        
	        $txtEmailRecipient.Enabled = $false
	        $txtEmailSender.Enabled = $false
	        $txtSMTPServer.Enabled = $false
		}
	}
	function OnApplicationLoad {
	    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
	    { 
	        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 
	    }
	    
		return $true #return true for success or false for failure
	}
	$chkEnterprise_CheckedChanged={
	    if ($chkEnterprise.Checked)
		{
			Load-ComboBox $comboDAG ''
	        Load-ComboBox $comboServer ''
	        Load-ComboBox $comboDatabase ''
	        $MailboxServerScope = @()
	        $comboDAG.Enabled = $false
	        $comboServer.Enabled = $false
	        $comboDatabase.Enabled = $false
	        #$datagridMailboxes
	     }
		else
		{
	        PopulateDAGs
	        PopulateMailboxServers
	        $comboDAG.Enabled = $true
	        $comboServer.Enabled = $true
	        $comboDatabase.Enabled = $false
	        $MailboxServerScope = @()
		}
	}
	
	$chkCurrentUser_CheckedChanged={
	    if ($chkCurrentUser.Checked)
		{
	        $txtUser.Enabled = $false
	        $txtPassword.Enabled = $false
		}
		else
		{
	        $txtUser.Enabled = $true
	        $txtPassword.Enabled = $true
		}
	}
	
	$buttonConnect_Click={
	    Get-PSSession | ?{$_.ComputerName -eq $txtServer.Text} | Remove-PSSession
	    If ($chkCurrentUser.Checked)
	    {
	        If (SnapinsAvailable) 
	        {
	            LoadSnapins
	            #. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1' Connect-ExchangeServer -auto
	            $EXConnected = $true
	            $MainForm.Text = "Exchange Wizard (Online)"
	        }
	        Else
	        {
	            $MainForm.Text = "Exchange Wizard (Offline - Snapin not available)"
	            $EXConnected = $false
			}
	
	    }
	    else
	    {
	        $Session = New-ExchangeSession -computername $txtServer.Text -user $txtUser.Text -pass $txtPassword.Text
	        #Import-PSSession $Session
	        $EXConnected = $true
	        $MainForm.Text = "Exchange Wizard (Online)"
	    }
	    if ($EXConnected)
	    {
	        $grpScope.Enabled = $true
		}
	}
	
	$comboServer_SelectedIndexChanged={
		if ($comboServer.Text -ne '')
	    {
	        $comboDatabase.Enabled = $true
	        PopulateDatabases
		}
	    else
	    {
	        $comboDatabase.Enabled = $false
	        Load-ComboBox $comboDatabase ''
		}
	}
	
	$comboDAG_SelectedIndexChanged={
		if ($comboDAG.Text -ne '')
	    {        
	        PopulateDAGMailboxServers        
		}
	}
	
	function Get-MailboxList 
	{ 
	    Begin 
	    { 
	    } 
	    Process 
	    { 
	        $MBoxes = @()
	        if ($datagridviewMailboxes.SelectedRows.Count -ge 1)
	        {
		        $datagridviewMailboxes.SelectedRows | ForEach {$MBoxes += $Mailboxes[$_.Index]}
			}
	        elseif ($chkEnterprise.Checked)
	        {
	            $MBoxes = @(Get-Mailbox -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	    	}
	        elseif ($comboDatabase.Text -ne '')
	        {
	            $MBoxes = @(Get-Mailbox -Database $comboDatabase.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)      
	        }
	        elseif ($comboServer.Text -ne '')
	        {
	            $MBoxes = @(Get-Mailbox -Server $comboServer.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        }
	        elseif ($comboDAG.Text -ne '')
	        {
	            foreach ($Server in $Script:MailboxServerScope)
	            {
	                $MBoxes += @(Get-Mailbox -Server $Server -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	    		}
	    	}
	    }  
	      
	    End 
	    { 
	        Return $MBoxes 
	    }  
	}
	
	$btnLoadMailboxes_Click={
		if ($chkEnterprise.Checked)
	    {
	        $Mailboxes = @(Get-Mailbox -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $Mailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh
		}
	    elseif ($comboDatabase.Text -ne '')
	    {
	        $Mailboxes = @(Get-Mailbox -Database $comboDatabase.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $Mailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh           
	    }
	    elseif ($comboServer.Text -ne '')
	    {
	        $Mailboxes = @(Get-Mailbox -Server $comboServer.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $Mailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh   
	    }
	    elseif ($comboDAG.Text -ne '')
	    {
	        $Mailboxes = @()
	        foreach ($Server in $Script:MailboxServerScope)
	        {
	            $Mailboxes += @(Get-Mailbox -Server $Server -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
			}
	        $MailboxList = $Mailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh
		}
	}
	
	$buttonExecute_Click={
	    Custom-Function (Get-MailboxList)
	    #[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	    [void][System.Windows.Forms.MessageBox]::Show("Custom Script Complete!","Done")
	}
	
		# --End User Generated Script--
	#----------------------------------------------
	#region Generated Events
	#----------------------------------------------
	
	$Form_StateCorrection_Load=
	{
		#Correct the initial state of the form to prevent the .Net maximized form issue
		$MainForm.WindowState = $InitialFormWindowState
	}
	
	$Form_StoreValues_Closing=
	{
		#Store the control values
		$script:MainForm_txtServer = $txtServer.Text
		$script:MainForm_chkCurrentUser = $chkCurrentUser.Checked
		$script:MainForm_txtUser = $txtUser.Text
		$script:MainForm_comboDatabase = $comboDatabase.Text
		$script:MainForm_comboServer = $comboServer.Text
		$script:MainForm_comboDAG = $comboDAG.Text
		$script:MainForm_chkEnterprise = $chkEnterprise.Checked
		$script:MainForm_richtextAbout = $richtextAbout.Text
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$buttonExecute.remove_Click($buttonExecute_Click)
			$buttonConnect.remove_Click($buttonConnect_Click)
			$chkCurrentUser.remove_CheckedChanged($chkCurrentUser_CheckedChanged)
			$btnLoadMailboxes.remove_Click($btnLoadMailboxes_Click)
			$comboServer.remove_SelectedIndexChanged($comboServer_SelectedIndexChanged)
			$comboDAG.remove_SelectedIndexChanged($comboDAG_SelectedIndexChanged)
			$chkEnterprise.remove_CheckedChanged($chkEnterprise_CheckedChanged)
			$MainForm.remove_Load($form1_FadeInLoad)
			$timerFadeIn.remove_Tick($timerFadeIn_Tick)
			$MainForm.remove_Load($Form_StateCorrection_Load)
			$MainForm.remove_Closing($Form_StoreValues_Closing)
			$MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	#endregion Generated Events

	#----------------------------------------------
	#region Generated Form Code
	#----------------------------------------------
	#
	# MainForm
	#
	$MainForm.Controls.Add($buttonExecute)
	$MainForm.Controls.Add($datagridviewMailboxes)
	$MainForm.Controls.Add($buttonCancel)
	$MainForm.Controls.Add($tabcontrol1)
	$MainForm.ClientSize = '837, 463'
	$MainForm.Name = "MainForm"
	$MainForm.StartPosition = 'CenterScreen'
	$MainForm.Text = "Exchange Mailbox Wizard (Not Connected)"
	$MainForm.add_Load($form1_FadeInLoad)
	#
	# buttonExecute
	#
	$buttonExecute.Location = '669, 432'
	$buttonExecute.Name = "buttonExecute"
	$buttonExecute.Size = '75, 23'
	$buttonExecute.TabIndex = 206
	$buttonExecute.Text = "Execute!"
	$buttonExecute.UseVisualStyleBackColor = $True
	$buttonExecute.add_Click($buttonExecute_Click)
	#
	# datagridviewMailboxes
	#
	$datagridviewMailboxes.AllowUserToAddRows = $False
	$datagridviewMailboxes.AllowUserToDeleteRows = $False
	$datagridviewMailboxes.ColumnHeadersHeightSizeMode = 'AutoSize'
	$datagridviewMailboxes.Location = '421, 22'
	$datagridviewMailboxes.Name = "datagridviewMailboxes"
	$datagridviewMailboxes.ReadOnly = $True
	$datagridviewMailboxes.SelectionMode = 'FullRowSelect'
	$datagridviewMailboxes.ShowEditingIcon = $False
	$datagridviewMailboxes.Size = '410, 404'
	$datagridviewMailboxes.TabIndex = 204
	#
	# buttonCancel
	#
	$buttonCancel.DialogResult = 'Cancel'
	$buttonCancel.Location = '750, 432'
	$buttonCancel.Name = "buttonCancel"
	$buttonCancel.Size = '75, 23'
	$buttonCancel.TabIndex = 202
	$buttonCancel.Text = "Cancel"
	$buttonCancel.UseVisualStyleBackColor = $True
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabScope)
	$tabcontrol1.Controls.Add($tabAbout)
	$tabcontrol1.Location = '-5, 0'
	$tabcontrol1.Name = "tabcontrol1"
	$tabcontrol1.SelectedIndex = 0
	$tabcontrol1.Size = '420, 430'
	$tabcontrol1.TabIndex = 40
	#
	# tabScope
	#
	$tabScope.Controls.Add($grpCredentials)
	$tabScope.Controls.Add($grpScope)
	$tabScope.BackColor = 'ControlLight'
	$tabScope.Location = '4, 22'
	$tabScope.Name = "tabScope"
	$tabScope.Padding = '3, 3, 3, 3'
	$tabScope.Size = '412, 404'
	$tabScope.TabIndex = 0
	$tabScope.Text = "Script Settings"
	#
	# grpCredentials
	#
	$grpCredentials.Controls.Add($txtServer)
	$grpCredentials.Controls.Add($labelServer)
	$grpCredentials.Controls.Add($buttonConnect)
	$grpCredentials.Controls.Add($chkCurrentUser)
	$grpCredentials.Controls.Add($labelCurrentUser)
	$grpCredentials.Controls.Add($txtUser)
	$grpCredentials.Controls.Add($txtPassword)
	$grpCredentials.Controls.Add($labelUser)
	$grpCredentials.Controls.Add($labelPassword)
	$grpCredentials.Location = '6, 6'
	$grpCredentials.Name = "grpCredentials"
	$grpCredentials.Size = '401, 110'
	$grpCredentials.TabIndex = 1
	$grpCredentials.TabStop = $False
	$grpCredentials.Text = "Credentials"
	#
	# txtServer
	#
	$txtServer.Location = '102, 15'
	$txtServer.Name = "txtServer"
	$txtServer.Size = '211, 20'
	$txtServer.TabIndex = 0
	$txtServer.Text = "localhost"
	#
	# labelServer
	#
	$labelServer.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelServer.Location = '35, 15'
	$labelServer.Name = "labelServer"
	$labelServer.Size = '51, 20'
	$labelServer.TabIndex = 88
	$labelServer.Text = "Server"
	$labelServer.TextAlign = 'MiddleRight'
	#
	# buttonConnect
	#
	$buttonConnect.Location = '320, 81'
	$buttonConnect.Name = "buttonConnect"
	$buttonConnect.Size = '73, 23'
	$buttonConnect.TabIndex = 4
	$buttonConnect.Text = "Connect!"
	$buttonConnect.UseVisualStyleBackColor = $True
	$buttonConnect.add_Click($buttonConnect_Click)
	#
	# chkCurrentUser
	#
	$chkCurrentUser.Checked = $True
	$chkCurrentUser.CheckState = 'Checked'
	$chkCurrentUser.Location = '102, 38'
	$chkCurrentUser.Name = "chkCurrentUser"
	$chkCurrentUser.Size = '16, 20'
	$chkCurrentUser.TabIndex = 1
	$chkCurrentUser.UseVisualStyleBackColor = $True
	$chkCurrentUser.add_CheckedChanged($chkCurrentUser_CheckedChanged)
	#
	# labelCurrentUser
	#
	$labelCurrentUser.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelCurrentUser.Location = '3, 38'
	$labelCurrentUser.Name = "labelCurrentUser"
	$labelCurrentUser.Size = '83, 20'
	$labelCurrentUser.TabIndex = 85
	$labelCurrentUser.Text = "Current User"
	$labelCurrentUser.TextAlign = 'MiddleRight'
	#
	# txtUser
	#
	$txtUser.Enabled = $False
	$txtUser.Location = '102, 58'
	$txtUser.Name = "txtUser"
	$txtUser.Size = '211, 20'
	$txtUser.TabIndex = 2
	#
	# txtPassword
	#
	$txtPassword.Enabled = $False
	$txtPassword.Location = '102, 84'
	$txtPassword.Name = "txtPassword"
	$txtPassword.PasswordChar = '*'
	$txtPassword.Size = '212, 20'
	$txtPassword.TabIndex = 3
	#
	# labelUser
	#
	$labelUser.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelUser.Location = '24, 59'
	$labelUser.Name = "labelUser"
	$labelUser.Size = '63, 21'
	$labelUser.TabIndex = 80
	$labelUser.Text = "User"
	$labelUser.TextAlign = 'MiddleRight'
	#
	# labelPassword
	#
	$labelPassword.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelPassword.Location = '23, 86'
	$labelPassword.Name = "labelPassword"
	$labelPassword.Size = '63, 20'
	$labelPassword.TabIndex = 81
	$labelPassword.Text = "Password"
	$labelPassword.TextAlign = 'MiddleRight'
	#
	# grpScope
	#
	$grpScope.Controls.Add($btnLoadMailboxes)
	$grpScope.Controls.Add($labelMailboxes)
	$grpScope.Controls.Add($comboDatabase)
	$grpScope.Controls.Add($labelDatabase)
	$grpScope.Controls.Add($comboServer)
	$grpScope.Controls.Add($comboDAG)
	$grpScope.Controls.Add($chkEnterprise)
	$grpScope.Controls.Add($labelCluster)
	$grpScope.Controls.Add($labelEnterprise)
	$grpScope.Controls.Add($labelDAG)
	$grpScope.Enabled = $False
	$grpScope.Location = '7, 122'
	$grpScope.Name = "grpScope"
	$grpScope.Size = '400, 276'
	$grpScope.TabIndex = 2
	$grpScope.TabStop = $False
	$grpScope.Text = "Scope"
	#
	# btnLoadMailboxes
	#
	$btnLoadMailboxes.Location = '153, 93'
	$btnLoadMailboxes.Name = "btnLoadMailboxes"
	$btnLoadMailboxes.Size = '75, 23'
	$btnLoadMailboxes.TabIndex = 92
	$btnLoadMailboxes.Text = "Load"
	$btnLoadMailboxes.UseVisualStyleBackColor = $True
	$btnLoadMailboxes.add_Click($btnLoadMailboxes_Click)
	#
	# labelMailboxes
	#
	$labelMailboxes.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelMailboxes.Location = '79, 94'
	$labelMailboxes.Name = "labelMailboxes"
	$labelMailboxes.Size = '73, 20'
	$labelMailboxes.TabIndex = 91
	$labelMailboxes.Text = "Mailboxes"
	$labelMailboxes.TextAlign = 'MiddleRight'
	#
	# comboDatabase
	#
	$comboDatabase.DropDownStyle = 'DropDownList'
	$comboDatabase.Enabled = $False
	$comboDatabase.FormattingEnabled = $True
	$comboDatabase.Location = '153, 66'
	$comboDatabase.Name = "comboDatabase"
	$comboDatabase.Size = '239, 21'
	$comboDatabase.Sorted = $True
	$comboDatabase.TabIndex = 8
	#
	# labelDatabase
	#
	$labelDatabase.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelDatabase.Location = '90, 67'
	$labelDatabase.Name = "labelDatabase"
	$labelDatabase.Size = '62, 20'
	$labelDatabase.TabIndex = 88
	$labelDatabase.Text = "Database"
	$labelDatabase.TextAlign = 'MiddleRight'
	#
	# comboServer
	#
	$comboServer.DropDownStyle = 'DropDownList'
	$comboServer.Enabled = $False
	$comboServer.FormattingEnabled = $True
	$comboServer.Location = '153, 39'
	$comboServer.Name = "comboServer"
	$comboServer.Size = '239, 21'
	$comboServer.TabIndex = 7
	$comboServer.add_SelectedIndexChanged($comboServer_SelectedIndexChanged)
	#
	# comboDAG
	#
	$comboDAG.DropDownStyle = 'DropDownList'
	$comboDAG.Enabled = $False
	$comboDAG.FormattingEnabled = $True
	$comboDAG.Location = '153, 12'
	$comboDAG.Name = "comboDAG"
	$comboDAG.Size = '239, 21'
	$comboDAG.TabIndex = 6
	$comboDAG.add_SelectedIndexChanged($comboDAG_SelectedIndexChanged)
	#
	# chkEnterprise
	#
	$chkEnterprise.Checked = $True
	$chkEnterprise.CheckState = 'Checked'
	$chkEnterprise.Location = '76, 12'
	$chkEnterprise.Name = "chkEnterprise"
	$chkEnterprise.Size = '19, 24'
	$chkEnterprise.TabIndex = 5
	$chkEnterprise.UseVisualStyleBackColor = $True
	$chkEnterprise.add_CheckedChanged($chkEnterprise_CheckedChanged)
	#
	# labelCluster
	#
	$labelCluster.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelCluster.Location = '101, 40'
	$labelCluster.Name = "labelCluster"
	$labelCluster.Size = '51, 20'
	$labelCluster.TabIndex = 83
	$labelCluster.Text = "Server"
	$labelCluster.TextAlign = 'MiddleRight'
	#
	# labelEnterprise
	#
	$labelEnterprise.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelEnterprise.Location = '6, 13'
	$labelEnterprise.Name = "labelEnterprise"
	$labelEnterprise.Size = '64, 20'
	$labelEnterprise.TabIndex = 82
	$labelEnterprise.Text = "Enterprise"
	$labelEnterprise.TextAlign = 'MiddleRight'
	#
	# labelDAG
	#
	$labelDAG.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelDAG.Location = '117, 12'
	$labelDAG.Name = "labelDAG"
	$labelDAG.Size = '35, 20'
	$labelDAG.TabIndex = 81
	$labelDAG.Text = "DAG"
	$labelDAG.TextAlign = 'MiddleRight'
	#
	# tabAbout
	#
	$tabAbout.Controls.Add($richtextAbout)
	$tabAbout.BackColor = 'ControlLight'
	$tabAbout.Location = '4, 22'
	$tabAbout.Name = "tabAbout"
	$tabAbout.Size = '412, 404'
	$tabAbout.TabIndex = 4
	$tabAbout.Text = "About"
	#
	# richtextAbout
	#
	$richtextAbout.Location = '3, 4'
	$richtextAbout.Name = "richtextAbout"
	$richtextAbout.ScrollBars = 'Vertical'
	$richtextAbout.Size = '406, 397'
	$richtextAbout.TabIndex = 0
	$richtextAbout.Text = "Exchange 2010 Mailbox GUI
Author: Zachary Loeber
Site: http://www.the-little-things.net/blog/2013/04/07/exchange-mailbox-gui/
Description: A powershell GUI for selecting and performing actions against multiple Exchange mailboxes.

"
	#
	# timerFadeIn
	#
	$timerFadeIn.add_Tick($timerFadeIn_Tick)
	#
	# tooltipAll
	#
	#endregion Generated Form Code

	#----------------------------------------------

	#Save the initial state of the form
	$InitialFormWindowState = $MainForm.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$MainForm.add_Load($Form_StateCorrection_Load)
	#Clean up the control events
	$MainForm.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$MainForm.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $MainForm.ShowDialog()

}
#endregion Source: MainForm.pff

#region Source: Globals.ps1
	#========================================================================
	# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.14
	# Created on:   2/5/2013 6:36 AM
	# Created by:   Zachary Loeber
	# Organization: 
	# Filename: Globals.ps1
	# Description: These are used across both the gui and the called script
	#              for storing and loading script state data.
	#========================================================================
	
	# Some dot sourcing love
	. .\Custom-Script.ps1
	
	# RequiredSnapins
	$Snapins=@(’Microsoft.Exchange.Management.PowerShell.E2010’)
	
	#Our base variables
	$varServer="localhost"
	$varUseCurrentUser=$true
	$varUser=""
	$varPassword=""
	$varScopeEnterprise=$true
	$varScopeDAG=""
	$varScopeServer=""
	$varScopeMailboxDatabase=""
	
	# Exchange specific globals that do not get saved
	$EXConnected = $false
	$MailboxServerScope = @()
	$Mailboxes = @()
	$SelectedMailboxes =@()
	
	#Provides the location of the script
	function Get-ScriptDirectory
	{ 
		if($hostinvocation -ne $null)
		{
			Split-Path $hostinvocation.MyCommand.path
		}
		else
		{
			Split-Path $script:MyInvocation.MyCommand.Path
		}
	}
	
	#Provides the location of the script
	[string]$ScriptDirectory = Get-ScriptDirectory
	
	function LoadSnapins
	{
	    $AllRequiredSnapinsLoaded=$True
	    if (($Snapins.Count -ge 1) -and $(SnapinsAvailable)) 
	    {
	    	Foreach ($Snapin in $Snapins)
	    	{
	            Add-PSSnapin $Snapin –ErrorAction SilentlyContinue 
	    		if ((Get-PSSnapin $Snapin –ErrorAction SilentlyContinue) –eq $NULL) 
	    		{
	    			$AllRequiredSnapinsLoaded=$false
	    		}
	     	}
	    }
	    else
	    {
	        $AllRequiredSnapinsLoaded=$false
	    }
	    Return $AllRequiredSnapinsLoaded
	}
	
	function SnapinsAvailable
	{
	    $RegisteredSnapins=@(Get-PSSnapin -Registered)
	    $RequiredSnapinsRegistered = $true
	    if ($Snapins.Count -ge 1) 
	    {
	    	Foreach ($Snapin in $Snapins)
	    	{
	            if (!($RegisteredSnapins -match $Snapin))
	            {
	                $RequiredSnapinsRegistered = $false
	            }
			}
		}
	    
	    Return $RequiredSnapinsRegistered
	}
	
	####################### 
	function Get-Type 
	{ 
	    param($type) 
	 
	$types = @( 
	'System.Boolean', 
	'System.Byte[]', 
	'System.Byte', 
	'System.Char', 
	'System.Datetime', 
	'System.Decimal', 
	'System.Double', 
	'System.Guid', 
	'System.Int16', 
	'System.Int32', 
	'System.Int64', 
	'System.Single', 
	'System.UInt16', 
	'System.UInt32', 
	'System.UInt64') 
	 
	    if ( $types -contains $type ) { 
	        Write-Output "$type" 
	    } 
	    else { 
	        Write-Output 'System.String' 
	         
	    } 
	} #Get-Type 
	 
	####################### 
	<# 
	.SYNOPSIS 
	Creates a DataTable for an object 
	.DESCRIPTION 
	Creates a DataTable based on an objects properties. 
	.INPUTS 
	Object 
	    Any object can be piped to Out-DataTable 
	.OUTPUTS 
	   System.Data.DataTable 
	.EXAMPLE 
	$dt = Get-psdrive| Out-DataTable 
	This example creates a DataTable from the properties of Get-psdrive and assigns output to $dt variable 
	.NOTES 
	Adapted from script by Marc van Orsouw see link 
	Version History 
	v1.0  - Chad Miller - Initial Release 
	v1.1  - Chad Miller - Fixed Issue with Properties 
	v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
	v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
	v1.4  - Chad Miller - Corrected issue with DBNull 
	v1.5  - Chad Miller - Updated example 
	v1.6  - Chad Miller - Added column datatype logic with default to string 
	v1.7 - Chad Miller - Fixed issue with IsArray 
	.LINK 
	http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx 
	#> 
	function Out-DataTable 
	{ 
	    [CmdletBinding()] 
	    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
	 
	    Begin 
	    { 
	        $dt = new-object Data.datatable   
	        $First = $true  
	    } 
	    Process 
	    { 
	        foreach ($object in $InputObject) 
	        { 
	            $DR = $DT.NewRow()   
	            foreach($property in $object.PsObject.get_properties()) 
	            {   
	                if ($first) 
	                {   
	                    $Col =  new-object Data.DataColumn   
	                    $Col.ColumnName = $property.Name.ToString()   
	                    if ($property.value) 
	                    { 
	                        if ($property.value -isnot [System.DBNull]) { 
	                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
	                         } 
	                    } 
	                    $DT.Columns.Add($Col) 
	                }   
	                if ($property.Gettype().IsArray) { 
	                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
	                }   
	               else { 
	                    $DR.Item($property.Name) = $property.value 
	                } 
	            }   
	            $DT.Rows.Add($DR)   
	            $First = $false 
	        } 
	    }  
	      
	    End 
	    { 
	        Write-Output @(,($dt)) 
	    } 
	 
	} #Out-DataTable
	
	Function New-ExchangeSession 
	{ 
	<# 
	   .Synopsis 
	    This function creates an implicit remoting connection to an Exchange Server  
	   .Description 
	    This function creates an implicit remoting session to a remote Exchange  
	    Server. It has been tested on Exchange 2010. The Exchange commands are 
	    brought into the local PowerShell environment. This works in both the 
	    Windows PowerShell console as well as the Windows PowerShell ISE. It requires 
	    two parameters: the computername and the user name with rights on the remote  
	    Exchange server. 
	   .Example 
	    New-ExchangeSession -computername ex1 -user iammred\administrator 
	    Makes an implicit remoting connection to a remote Exchange 2010 server 
	    named ex1 using the administrator account from the iammred domain. The user 
	    is prompted for the administrator password. 
	   .Parameter ComputerName 
	    The name of the remote Exchange server 
	   .Parameter User 
	    The user account with rights on the remote Exchange server. The user 
	    account is specified as domain\username 
	   .Notes 
	    NAME:  New-ExchangeSession 
	    AUTHOR: ed wilson, msft 
	    LASTEDIT: 01/13/2012 17:05:32 
	    KEYWORDS: Messaging & Communication, Microsoft Exchange 2010, Remoting 
	    HSG: HSG-1-23-12 
	   .Link 
	     Http://www.ScriptingGuys.com 
	 #Requires -Version 2.0 
	 #> 
	     Param( 
	      [Parameter(Mandatory=$true,Position=0)] 
	      [String] 
	      $computername, 
	      [Parameter(Mandatory=$true,Position=1)] 
	      [String] 
	      $user,
	      [Parameter(Mandatory=$true,Position=2)] 
	      [String] 
	      $pass
	      ) 
	    $pass2 = ConvertTo-SecureString -AsPlainText $pass -Force
	    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $user,$pass2
	    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$computername/powershell -Credential $credential
	    Import-PSSession $session
	} #end function New-ExchangeSession
#endregion Source: Globals.ps1

#Start the application
Main ($CommandLine)

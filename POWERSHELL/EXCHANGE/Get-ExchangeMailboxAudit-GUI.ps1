#region Source: Startup.pfs
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
	if((Call-MainForm_pff) -eq "OK") {
		
	}
	
	$global:ExitCode = 0 #Set the exit code for the Packager
}
#endregion Source: Startup.pfs

#region Source: MainForm.pff
function Call-MainForm_pff
{
#region File Recovery Data (DO NOT MODIFY)

#endregion
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
	$buttonSave = New-Object 'System.Windows.Forms.Button'
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
	$tabOptions = New-Object 'System.Windows.Forms.TabPage'
	$grpMailboxPermissions = New-Object 'System.Windows.Forms.GroupBox'
	$groupboxFlagging = New-Object 'System.Windows.Forms.GroupBox'
	$labelDeletedItemTotalSize = New-Object 'System.Windows.Forms.Label'
	$dialDeletedSizeAlert = New-Object 'System.Windows.Forms.NumericUpDown'
	$label2 = New-Object 'System.Windows.Forms.Label'
	$label3 = New-Object 'System.Windows.Forms.Label'
	$dialDeletedSizeWarn = New-Object 'System.Windows.Forms.NumericUpDown'
	$labelMailboxTotalSizeMB = New-Object 'System.Windows.Forms.Label'
	$dialTotalSizeAlert = New-Object 'System.Windows.Forms.NumericUpDown'
	$labelAlert = New-Object 'System.Windows.Forms.Label'
	$labelWarning = New-Object 'System.Windows.Forms.Label'
	$dialTotalSizeWarn = New-Object 'System.Windows.Forms.NumericUpDown'
	$grpMailPermsGeneral = New-Object 'System.Windows.Forms.GroupBox'
	$chkFlagWarnings = New-Object 'System.Windows.Forms.CheckBox'
	$chkExcludeUnknown = New-Object 'System.Windows.Forms.CheckBox'
	$chkExcludeZeroResults = New-Object 'System.Windows.Forms.CheckBox'
	$chkIncludeInherited = New-Object 'System.Windows.Forms.CheckBox'
	$buttonRemoveIgnored = New-Object 'System.Windows.Forms.Button'
	$buttonAddIgnored = New-Object 'System.Windows.Forms.Button'
	$txtPermIgnore = New-Object 'System.Windows.Forms.TextBox'
	$labelIgnoredUsers = New-Object 'System.Windows.Forms.Label'
	$labelIncludedReportElemen = New-Object 'System.Windows.Forms.Label'
	$listPermReportIgnore = New-Object 'System.Windows.Forms.ListBox'
	$listPermReportOptions = New-Object 'System.Windows.Forms.ListBox'
	$tabReportDeliveryOptions = New-Object 'System.Windows.Forms.TabPage'
	$grpReportFormat = New-Object 'System.Windows.Forms.GroupBox'
	$labelReportStyle = New-Object 'System.Windows.Forms.Label'
	$comboReportStyle = New-Object 'System.Windows.Forms.ComboBox'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$txtEmailSubject = New-Object 'System.Windows.Forms.TextBox'
	$labelSubject = New-Object 'System.Windows.Forms.Label'
	$labelEmailReport = New-Object 'System.Windows.Forms.Label'
	$chkEmailReport = New-Object 'System.Windows.Forms.CheckBox'
	$labelReportName = New-Object 'System.Windows.Forms.Label'
	$labelSaveReport = New-Object 'System.Windows.Forms.Label'
	$txtReportName = New-Object 'System.Windows.Forms.TextBox'
	$labelReportFolder = New-Object 'System.Windows.Forms.Label'
	$chkSaveLocally = New-Object 'System.Windows.Forms.CheckBox'
	$txtReportFolder = New-Object 'System.Windows.Forms.TextBox'
	$txtSMTPServer = New-Object 'System.Windows.Forms.TextBox'
	$buttonBrowseFolder = New-Object 'System.Windows.Forms.Button'
	$labelSMTPRelayServer = New-Object 'System.Windows.Forms.Label'
	$labelEmailSender = New-Object 'System.Windows.Forms.Label'
	$labelEmailRecipient = New-Object 'System.Windows.Forms.Label'
	$txtEmailRecipient = New-Object 'System.Windows.Forms.TextBox'
	$txtEmailSender = New-Object 'System.Windows.Forms.TextBox'
	$tabAbout = New-Object 'System.Windows.Forms.TabPage'
	$richtextboxAbout1 = New-Object 'System.Windows.Forms.RichTextBox'
	$timerFadeIn = New-Object 'System.Windows.Forms.Timer'
	$folderbrowserdialog1 = New-Object 'System.Windows.Forms.FolderBrowserDialog'
	$tooltipAll = New-Object 'System.Windows.Forms.ToolTip'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#endregion Generated Form Objects

	#----------------------------------------------
	# User Generated Script
	#----------------------------------------------
	Function Set-SaveData
	{
	    $Script:varEmailReport = $chkEmailReport.Checked
	    $Script:varEmailSubject = $txtEmailSubject.Text
	    $Script:varEmailRecipient = $txtEmailRecipient.Text
	    $Script:varEmailSender = $txtEmailSender.Text
	    $Script:varSMTPServer = $txtSMTPServer.Text
	    $Script:varSaveReportsLocally = $chkSaveLocally.Checked
	    $Script:varReportName = $txtReportName.Text
	    $Script:varReportFolder = $txtReportFolder.Text
	    $Script:varServer = $txtServer.Text
	    $Script:varUseCurrentUser = $chkCurrentUser.Checked
	    $Script:varUser = $txtUser.Text
		$Script:varScopeEnterprise = $chkEnterprise.Checked
		$Script:varScopeDAG = $comboDAG.Text
		$Script:varScopeServer = $comboServer.Text
		$Script:varScopeMailboxDatabase= $comboDatabase.Text
		$Script:varIncludeInherited = $chkIncludeInherited.Checked
		$Script:varExcludeZeroResults = $chkExcludeZeroResults.Checked
		$Script:varExcludeUnknownUsers = $chkExcludeUnknown.Checked
		$Script:varSummaryReport = $false
		$Script:varFullAccessReport = $false
		$Script:varSendAsReport = $false
		$Script:varSendOnBehalfReport = $false
		$Script:varCalendarPermReport = $false
		foreach ($SelectedReportOption in $listPermReportOptions.SelectedItems)
	    {
			if ($SelectedReportOption.Option -eq 'Mailbox Summary Information')  {$Script:varSummaryReport = $true}
			if ($SelectedReportOption.Option -eq 'Full Access Permissions')  {$Script:varFullAccessReport = $true}
			if ($SelectedReportOption.Option -eq 'Send On Behalf Permissions')  {$Script:varSendAsReport = $true}
			if ($SelectedReportOption.Option -eq 'Send As Permissions')  {$Script:varSendOnBehalfReport = $true}
			if ($SelectedReportOption.Option -eq 'Calendar Permissions')  {$Script:varCalendarPermReport = $true}
	        if ($SelectedReportOption.Option -eq 'Mailbox Rule - Forwarding')  {$Script:varMailboxRuleForwarding = $true}
	        if ($SelectedReportOption.Option -eq 'Mailbox Rule - Redirecting')  {$Script:varMailboxRuleRedirecting = $true}
		}
		$Script:varFlagWarnings = $chkFlagWarnings.Checked
		$Script:varMailboxSizeWarning = $dialTotalSizeWarn.Text
		$Script:varMailboxSizeAlert = $dialTotalSizeAlert.Text
		$Script:varDeletedSizeWarning = $dialDeletedSizeWarn.Text
		$Script:varDeletedSizeAlert = $dialDeletedSizeAlert.Text
	
	    $MboxPermissionsSelected = @()
	    $MboxIgnoredUsers = @()
	    foreach ($username in $listPermReportIgnore.Items)
	    {    
	        $MboxIgnoredUsers += $username
		}
	    
	    $MboxPermissionsSelected = @()
	    for($counter = 0; $counter -lt $listPermReportOptions.Items.Count; $counter++) 
	    {
	        $newobj = New-Object -TypeName PSObject
	        $newobj | Add-Member -Name Option -Value ($listPermReportOptions.Items[$counter]).Option -MemberType NoteProperty
	        if ($listPermReportOptions.SelectedItems -contains $listPermReportOptions.Items[$counter])
	        {
	            $newobj | Add-Member -Name Selected -Value $true -MemberType NoteProperty
			}
	        else
	        {
	            $newobj | Add-Member -Name Selected -Value $false -MemberType NoteProperty
			}
	        $MboxPermissionsSelected += $newobj
		}
	    $Script:varMailboxReportIgnoredUsers = $MboxIgnoredUsers
	    $Script:varMailboxReportPermissions = $MboxPermissionsSelected
	}
	
	function Load-FormConfig
	{
	    $chkEmailReport.Checked = $varEmailReport
	    $txtEmailSubject.Text = $varEmailSubject
	    $txtEmailRecipient.Text = $varEmailRecipient
	    $txtEmailSender.Text = $varEmailSender
	    $txtSMTPServer.Text = $varSMTPServer
	    $chkSaveLocally.Checked = $varSaveReportsLocally
	    $txtReportName.Text = $varReportName
	    $txtReportFolder.Text = $varReportFolder
	    $txtServer.Text = $varServer
	    $chkCurrentUser.Checked = $varUseCurrentUser
	    $txtUser.Text = $varUser
		$chkIncludeInherited.Checked = $varIncludeInherited
		$chkExcludeZeroResults.Checked = $varExcludeZeroResults
		$chkExcludeUnknown.Checked = $varExcludeUnknownUsers
		$chkFlagWarnings.Checked = $varFlagWarnings
		$dialTotalSizeWarn.Text = $varMailboxSizeWarning
		$dialTotalSizeAlert.Text = $varMailboxSizeAlert 
		$dialDeletedSizeWarn.Text = $varDeletedSizeWarning
		$dialDeletedSizeAlert.Text = $varDeletedSizeAlert
		
		Load-ListBox  $listPermReportOptions $varMailboxReportPermissions Option
	    Load-ListBox $listPermReportIgnore $varMailboxReportIgnoredUsers
	
		$counter = 0 
		foreach ($Option in $varMailboxReportPermissions)
	    {
	        if ($Option.Selected)
			{
				$listPermReportOptions.SetSelected($counter,$true)
			}		
			$counter++
		}
		
	  
	}
	
	# Account for all of our form control depenencies.
	function Set-FormControlsState
	{
	#    #Saved Locally Checkbox
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
	    
	    # Email Checkbox
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
	
		Load-Config
		Load-FormConfig
		Set-FormControlsState
		. .\New-ExchangeMailboxAudit.ps1
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
	
	$buttonBrowse_Click={
	
		if($openfiledialog1.ShowDialog() -eq 'OK')
		{
			$txtReportFolder.Text = $openfiledialog1.FileName
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
	        $grpMailboxPermissions.Enabled = $true
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
	
	$chkMailboxPermissionRep_CheckedChanged={
		if ($chkMailboxPermissionRep.Checked)
	    {
	        $grpMailboxPermissions.Enabled = $true
		}
	    else
	    {
	        $grpMailboxPermissions.Enabled = $false
		}
	}
	
	$comboDAG_SelectedIndexChanged={
		if ($comboDAG.Text -ne '')
	    {        
	        PopulateDAGMailboxServers        
		}
	}
	
	function Get-SelectedMailboxes
	{
	    Process 
	    { 
	        $MBoxes = @()
	        if ($datagridviewMailboxes.SelectedRows.Count -ge 1)
	        {
		        $datagridviewMailboxes.SelectedRows | ForEach {$MBoxes += $datagridMailboxes[$_.Index].Identity}
			}
			else
			{
	            $MBoxes = @(Get-MailboxList -WholeEnterprise $chkEnterprise.Checked `
	                                        -DatabaseName $comboDatabase.Text `
	                                        -ServerName $comboServer.Text `
	                                        -DAGName $comboDAG.Text)
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
	        $datagridMailboxes = @(Get-Mailbox -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $datagridMailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh
		}
	    elseif ($comboDatabase.Text -ne '')
	    {
	        $datagridMailboxes = @(Get-Mailbox -Database $comboDatabase.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $datagridMailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh           
	    }
	    elseif ($comboServer.Text -ne '')
	    {
	        $datagridMailboxes = @(Get-Mailbox -Server $comboServer.Text -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
	        $MailboxList = $datagridMailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh   
	    }
	    elseif ($comboDAG.Text -ne '')
	    {
	        $datagridMailboxes = @()
	        foreach ($Server in $Script:MailboxServerScope)
	        {
	            $datagridMailboxes += @(Get-Mailbox -Server $Server -ResultSize Unlimited | Select Alias,DisplayName,RecipientTypeDetails,PrimarySMTPAddress,Servername,Database,Identity | Sort-Object DisplayName)
			}
	        $MailboxList = $datagridMailboxes | Out-DataTable
	        $datagridviewMailboxes.DataSource = $MailboxList
	        $datagridviewMailboxes.Refresh
		}
	}
	
	$buttonExecute_Click={
	    Set-SaveData
	    if (Save-Config)
	    {
	        $MailboxesToProcess = @(Get-SelectedMailboxes)
	        New-ExchangeMailboxAuditReport -LoadConfiguration -Mailboxes $MailboxesToProcess
	        #[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	        [void][System.Windows.Forms.MessageBox]::Show("Script Complete!","Done")
		}
	    
	}
	$buttonAddIgnored_Click={
	    if ($txtPermIgnore.Text -ne '')
	    {
		    Load-ListBox $listPermReportIgnore $txtPermIgnore.Text -Append
		}
	}
	
	$buttonRemoveIgnored_Click={
	    if ($listPermReportIgnore.SelectedItems.Count -gt 0) 
	    {
	        $txtPermIgnore.Text = $listPermReportIgnore.SelectedItem
	        $listPermReportIgnore.Items.remove($listPermReportIgnore.SelectedItem)
	        $listPermReportIgnore.Refresh
	    }
	}
	$chkFlagWarnings_CheckedChanged={
		if ($chkFlagWarnings.Checked)
		{
			$groupboxFlagging.Enabled = $true
		}
		else
		{
			$groupboxFlagging.Enabled = $false
		}
		
	}
	$buttonBrowse_Click3={
	
		if($openfiledialog2.ShowDialog() -eq 'OK')
		{
			$textboxFile.Text = $openfiledialog2.FileName
		}
	}
	
	$buttonBrowse_Click4={
	
		if($openfiledialog2.ShowDialog() -eq 'OK')
		{
			$textboxFile.Text = $openfiledialog2.FileName
		}
	}
	
	$buttonBrowse_Click5={
	
		if($openfiledialog2.ShowDialog() -eq 'OK')
		{
			$textboxFile.Text = $openfiledialog2.FileName
		}
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
		$script:MainForm_dialDeletedSizeAlert = $dialDeletedSizeAlert.Value
		$script:MainForm_dialDeletedSizeWarn = $dialDeletedSizeWarn.Value
		$script:MainForm_dialTotalSizeAlert = $dialTotalSizeAlert.Value
		$script:MainForm_dialTotalSizeWarn = $dialTotalSizeWarn.Value
		$script:MainForm_chkFlagWarnings = $chkFlagWarnings.Checked
		$script:MainForm_chkExcludeUnknown = $chkExcludeUnknown.Checked
		$script:MainForm_chkExcludeZeroResults = $chkExcludeZeroResults.Checked
		$script:MainForm_chkIncludeInherited = $chkIncludeInherited.Checked
		$script:MainForm_txtPermIgnore = $txtPermIgnore.Text
		$script:MainForm_listPermReportIgnore = $listPermReportIgnore.SelectedItems
		$script:MainForm_listPermReportOptions = $listPermReportOptions.SelectedItems
		$script:MainForm_comboReportStyle = $comboReportStyle.Text
		$script:MainForm_txtEmailSubject = $txtEmailSubject.Text
		$script:MainForm_chkEmailReport = $chkEmailReport.Checked
		$script:MainForm_txtReportName = $txtReportName.Text
		$script:MainForm_chkSaveLocally = $chkSaveLocally.Checked
		$script:MainForm_txtReportFolder = $txtReportFolder.Text
		$script:MainForm_txtSMTPServer = $txtSMTPServer.Text
		$script:MainForm_txtEmailRecipient = $txtEmailRecipient.Text
		$script:MainForm_txtEmailSender = $txtEmailSender.Text
		$script:MainForm_richtextboxAbout1 = $richtextboxAbout1.Text
	}

	
	$Form_Cleanup_FormClosed=
	{
		#Remove all event handlers from the controls
		try
		{
			$buttonSave.remove_Click($buttonSave_Click)
			$buttonExecute.remove_Click($buttonExecute_Click)
			$buttonConnect.remove_Click($buttonConnect_Click)
			$chkCurrentUser.remove_CheckedChanged($chkCurrentUser_CheckedChanged)
			$btnLoadMailboxes.remove_Click($btnLoadMailboxes_Click)
			$comboServer.remove_SelectedIndexChanged($comboServer_SelectedIndexChanged)
			$comboDAG.remove_SelectedIndexChanged($comboDAG_SelectedIndexChanged)
			$chkEnterprise.remove_CheckedChanged($chkEnterprise_CheckedChanged)
			$chkFlagWarnings.remove_CheckedChanged($chkFlagWarnings_CheckedChanged)
			$buttonRemoveIgnored.remove_Click($buttonRemoveIgnored_Click)
			$buttonAddIgnored.remove_Click($buttonAddIgnored_Click)
			$chkEmailReport.remove_CheckedChanged($chkEmailReport_CheckedChanged)
			$chkSaveLocally.remove_CheckedChanged($chkSaveLocally_CheckedChanged)
			$buttonBrowseFolder.remove_Click($buttonBrowseFolder_Click)
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
	$MainForm.Controls.Add($buttonSave)
	$MainForm.Controls.Add($buttonExecute)
	$MainForm.Controls.Add($datagridviewMailboxes)
	$MainForm.Controls.Add($buttonCancel)
	$MainForm.Controls.Add($tabcontrol1)
	$MainForm.ClientSize = '837, 456'
	$MainForm.Name = "MainForm"
	$MainForm.StartPosition = 'CenterScreen'
	$MainForm.Text = "Exchange Mailbox Wizard (Not Connected)"
	$MainForm.add_Load($form1_FadeInLoad)
	#
	# buttonSave
	#
	$buttonSave.Location = '340, 432'
	$buttonSave.Name = "buttonSave"
	$buttonSave.Size = '75, 23'
	$buttonSave.TabIndex = 208
	$buttonSave.Text = "Save"
	$buttonSave.UseVisualStyleBackColor = $True
	$buttonSave.add_Click($buttonSave_Click)
	#
	# buttonExecute
	#
	$buttonExecute.Location = '675, 432'
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
	$System_Windows_Forms_DataGridViewCellStyle_1 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
	$System_Windows_Forms_DataGridViewCellStyle_1.Alignment = 'MiddleLeft'
	$System_Windows_Forms_DataGridViewCellStyle_1.BackColor = 'Control'
	$System_Windows_Forms_DataGridViewCellStyle_1.Font = "Microsoft Sans Serif, 8.25pt"
	$System_Windows_Forms_DataGridViewCellStyle_1.ForeColor = 'WindowText'
	$System_Windows_Forms_DataGridViewCellStyle_1.SelectionBackColor = 'Highlight'
	$System_Windows_Forms_DataGridViewCellStyle_1.SelectionForeColor = 'HighlightText'
	$System_Windows_Forms_DataGridViewCellStyle_1.WrapMode = 'True'
	$datagridviewMailboxes.ColumnHeadersDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_1
	$datagridviewMailboxes.ColumnHeadersHeightSizeMode = 'AutoSize'
	$System_Windows_Forms_DataGridViewCellStyle_2 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
	$System_Windows_Forms_DataGridViewCellStyle_2.Alignment = 'MiddleLeft'
	$System_Windows_Forms_DataGridViewCellStyle_2.BackColor = 'Window'
	$System_Windows_Forms_DataGridViewCellStyle_2.Font = "Microsoft Sans Serif, 8.25pt"
	$System_Windows_Forms_DataGridViewCellStyle_2.ForeColor = 'ControlText'
	$System_Windows_Forms_DataGridViewCellStyle_2.SelectionBackColor = 'Highlight'
	$System_Windows_Forms_DataGridViewCellStyle_2.SelectionForeColor = 'HighlightText'
	$System_Windows_Forms_DataGridViewCellStyle_2.WrapMode = 'False'
	$datagridviewMailboxes.DefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_2
	$datagridviewMailboxes.Location = '421, 22'
	$datagridviewMailboxes.Name = "datagridviewMailboxes"
	$datagridviewMailboxes.ReadOnly = $True
	$System_Windows_Forms_DataGridViewCellStyle_3 = New-Object 'System.Windows.Forms.DataGridViewCellStyle'
	$System_Windows_Forms_DataGridViewCellStyle_3.Alignment = 'MiddleLeft'
	$System_Windows_Forms_DataGridViewCellStyle_3.BackColor = 'Control'
	$System_Windows_Forms_DataGridViewCellStyle_3.Font = "Microsoft Sans Serif, 8.25pt"
	$System_Windows_Forms_DataGridViewCellStyle_3.ForeColor = 'WindowText'
	$System_Windows_Forms_DataGridViewCellStyle_3.SelectionBackColor = 'Highlight'
	$System_Windows_Forms_DataGridViewCellStyle_3.SelectionForeColor = 'HighlightText'
	$System_Windows_Forms_DataGridViewCellStyle_3.WrapMode = 'True'
	$datagridviewMailboxes.RowHeadersDefaultCellStyle = $System_Windows_Forms_DataGridViewCellStyle_3
	$datagridviewMailboxes.SelectionMode = 'FullRowSelect'
	$datagridviewMailboxes.ShowEditingIcon = $False
	$datagridviewMailboxes.Size = '410, 408'
	$datagridviewMailboxes.TabIndex = 204
	#
	# buttonCancel
	#
	$buttonCancel.DialogResult = 'Cancel'
	$buttonCancel.Location = '756, 432'
	$buttonCancel.Name = "buttonCancel"
	$buttonCancel.Size = '75, 23'
	$buttonCancel.TabIndex = 202
	$buttonCancel.Text = "Cancel"
	$buttonCancel.UseVisualStyleBackColor = $True
	#
	# tabcontrol1
	#
	$tabcontrol1.Controls.Add($tabScope)
	$tabcontrol1.Controls.Add($tabOptions)
	$tabcontrol1.Controls.Add($tabReportDeliveryOptions)
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
	$grpScope.Location = '6, 122'
	$grpScope.Name = "grpScope"
	$grpScope.Size = '400, 122'
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
	# tabOptions
	#
	$tabOptions.Controls.Add($grpMailboxPermissions)
	$tabOptions.BackColor = 'ControlLight'
	$tabOptions.Location = '4, 22'
	$tabOptions.Name = "tabOptions"
	$tabOptions.Size = '412, 404'
	$tabOptions.TabIndex = 5
	$tabOptions.Text = "Options"
	#
	# grpMailboxPermissions
	#
	$grpMailboxPermissions.Controls.Add($groupboxFlagging)
	$grpMailboxPermissions.Controls.Add($grpMailPermsGeneral)
	$grpMailboxPermissions.Controls.Add($buttonRemoveIgnored)
	$grpMailboxPermissions.Controls.Add($buttonAddIgnored)
	$grpMailboxPermissions.Controls.Add($txtPermIgnore)
	$grpMailboxPermissions.Controls.Add($labelIgnoredUsers)
	$grpMailboxPermissions.Controls.Add($labelIncludedReportElemen)
	$grpMailboxPermissions.Controls.Add($listPermReportIgnore)
	$grpMailboxPermissions.Controls.Add($listPermReportOptions)
	$grpMailboxPermissions.Location = '3, 19'
	$grpMailboxPermissions.Name = "grpMailboxPermissions"
	$grpMailboxPermissions.Size = '406, 382'
	$grpMailboxPermissions.TabIndex = 0
	$grpMailboxPermissions.TabStop = $False
	$grpMailboxPermissions.Text = "Mailbox Permission Report"
	#
	# groupboxFlagging
	#
	$groupboxFlagging.Controls.Add($labelDeletedItemTotalSize)
	$groupboxFlagging.Controls.Add($dialDeletedSizeAlert)
	$groupboxFlagging.Controls.Add($label2)
	$groupboxFlagging.Controls.Add($label3)
	$groupboxFlagging.Controls.Add($dialDeletedSizeWarn)
	$groupboxFlagging.Controls.Add($labelMailboxTotalSizeMB)
	$groupboxFlagging.Controls.Add($dialTotalSizeAlert)
	$groupboxFlagging.Controls.Add($labelAlert)
	$groupboxFlagging.Controls.Add($labelWarning)
	$groupboxFlagging.Controls.Add($dialTotalSizeWarn)
	$groupboxFlagging.Location = '197, 169'
	$groupboxFlagging.Name = "groupboxFlagging"
	$groupboxFlagging.Size = '202, 161'
	$groupboxFlagging.TabIndex = 7
	$groupboxFlagging.TabStop = $False
	$groupboxFlagging.Text = "Flagging"
	#
	# labelDeletedItemTotalSize
	#
	$labelDeletedItemTotalSize.Font = "Microsoft Sans Serif, 8.25pt, style=Bold, Underline"
	$labelDeletedItemTotalSize.Location = '6, 82'
	$labelDeletedItemTotalSize.Name = "labelDeletedItemTotalSize"
	$labelDeletedItemTotalSize.Size = '190, 19'
	$labelDeletedItemTotalSize.TabIndex = 16
	$labelDeletedItemTotalSize.Text = "Deleted Item Total Size (MB)"
	#
	# dialDeletedSizeAlert
	#
	$dialDeletedSizeAlert.Location = '149, 131'
	$dialDeletedSizeAlert.Maximum = 100000
	$dialDeletedSizeAlert.Name = "dialDeletedSizeAlert"
	$dialDeletedSizeAlert.Size = '47, 20'
	$dialDeletedSizeAlert.TabIndex = 3
	$dialDeletedSizeAlert.Value = 1024
	#
	# label2
	#
	$label2.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$label2.Location = '6, 133'
	$label2.Name = "label2"
	$label2.Size = '73, 18'
	$label2.TabIndex = 14
	$label2.Text = "Alert"
	#
	# label3
	#
	$label3.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$label3.Location = '6, 106'
	$label3.Name = "label3"
	$label3.Size = '73, 14'
	$label3.TabIndex = 13
	$label3.Text = "Warning"
	#
	# dialDeletedSizeWarn
	#
	$dialDeletedSizeWarn.Location = '149, 104'
	$dialDeletedSizeWarn.Maximum = 100000
	$dialDeletedSizeWarn.Name = "dialDeletedSizeWarn"
	$dialDeletedSizeWarn.Size = '47, 20'
	$dialDeletedSizeWarn.TabIndex = 2
	$dialDeletedSizeWarn.Value = 512
	#
	# labelMailboxTotalSizeMB
	#
	$labelMailboxTotalSizeMB.Font = "Microsoft Sans Serif, 8.25pt, style=Bold, Underline"
	$labelMailboxTotalSizeMB.Location = '6, 16'
	$labelMailboxTotalSizeMB.Name = "labelMailboxTotalSizeMB"
	$labelMailboxTotalSizeMB.Size = '148, 19'
	$labelMailboxTotalSizeMB.TabIndex = 11
	$labelMailboxTotalSizeMB.Text = "Mailbox Total Size (MB)"
	#
	# dialTotalSizeAlert
	#
	$dialTotalSizeAlert.Location = '149, 59'
	$dialTotalSizeAlert.Maximum = 100000
	$dialTotalSizeAlert.Name = "dialTotalSizeAlert"
	$dialTotalSizeAlert.Size = '47, 20'
	$dialTotalSizeAlert.TabIndex = 1
	$dialTotalSizeAlert.Value = 1024
	#
	# labelAlert
	#
	$labelAlert.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelAlert.Location = '6, 62'
	$labelAlert.Name = "labelAlert"
	$labelAlert.Size = '73, 20'
	$labelAlert.TabIndex = 9
	$labelAlert.Text = "Alert"
	#
	# labelWarning
	#
	$labelWarning.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelWarning.Location = '6, 37'
	$labelWarning.Name = "labelWarning"
	$labelWarning.Size = '73, 17'
	$labelWarning.TabIndex = 8
	$labelWarning.Text = "Warning"
	#
	# dialTotalSizeWarn
	#
	$dialTotalSizeWarn.Location = '149, 34'
	$dialTotalSizeWarn.Maximum = 100000
	$dialTotalSizeWarn.Name = "dialTotalSizeWarn"
	$dialTotalSizeWarn.Size = '47, 20'
	$dialTotalSizeWarn.TabIndex = 0
	$dialTotalSizeWarn.Value = 512
	#
	# grpMailPermsGeneral
	#
	$grpMailPermsGeneral.Controls.Add($chkFlagWarnings)
	$grpMailPermsGeneral.Controls.Add($chkExcludeUnknown)
	$grpMailPermsGeneral.Controls.Add($chkExcludeZeroResults)
	$grpMailPermsGeneral.Controls.Add($chkIncludeInherited)
	$grpMailPermsGeneral.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$grpMailPermsGeneral.Location = '6, 169'
	$grpMailPermsGeneral.Name = "grpMailPermsGeneral"
	$grpMailPermsGeneral.Size = '182, 110'
	$grpMailPermsGeneral.TabIndex = 5
	$grpMailPermsGeneral.TabStop = $False
	$grpMailPermsGeneral.Text = "General"
	#
	# chkFlagWarnings
	#
	$chkFlagWarnings.Checked = $True
	$chkFlagWarnings.CheckState = 'Checked'
	$chkFlagWarnings.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$chkFlagWarnings.Location = '5, 86'
	$chkFlagWarnings.Name = "chkFlagWarnings"
	$chkFlagWarnings.Size = '125, 18'
	$chkFlagWarnings.TabIndex = 4
	$chkFlagWarnings.Text = "Flag Warnings"
	$chkFlagWarnings.UseVisualStyleBackColor = $True
	$chkFlagWarnings.add_CheckedChanged($chkFlagWarnings_CheckedChanged)
	#
	# chkExcludeUnknown
	#
	$chkExcludeUnknown.Checked = $True
	$chkExcludeUnknown.CheckState = 'Checked'
	$chkExcludeUnknown.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$chkExcludeUnknown.Location = '5, 59'
	$chkExcludeUnknown.Name = "chkExcludeUnknown"
	$chkExcludeUnknown.Size = '171, 24'
	$chkExcludeUnknown.TabIndex = 3
	$chkExcludeUnknown.Text = "Exclude Unknown Users"
	$tooltipAll.SetToolTip($chkExcludeUnknown, "Ignore any user matching ""S-1-*"". These are the crazy looking accounts which don't seem to resolve (usually just deleted in AD).")
	$chkExcludeUnknown.UseVisualStyleBackColor = $True
	#
	# chkExcludeZeroResults
	#
	$chkExcludeZeroResults.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$chkExcludeZeroResults.Location = '5, 34'
	$chkExcludeZeroResults.Name = "chkExcludeZeroResults"
	$chkExcludeZeroResults.Size = '157, 29'
	$chkExcludeZeroResults.TabIndex = 2
	$chkExcludeZeroResults.Text = "Exclude Zero Results"
	$tooltipAll.SetToolTip($chkExcludeZeroResults, "If a mailbox results in nothing being found this will exclude it entirely from the report")
	$chkExcludeZeroResults.UseVisualStyleBackColor = $True
	#
	# chkIncludeInherited
	#
	$chkIncludeInherited.Checked = $True
	$chkIncludeInherited.CheckState = 'Checked'
	$chkIncludeInherited.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$chkIncludeInherited.Location = '6, 13'
	$chkIncludeInherited.Name = "chkIncludeInherited"
	$chkIncludeInherited.Size = '141, 24'
	$chkIncludeInherited.TabIndex = 1
	$chkIncludeInherited.Text = "Include Inherited"
	$tooltipAll.SetToolTip($chkIncludeInherited, "When performing permissions reports this option will include all those which are inherited")
	$chkIncludeInherited.UseVisualStyleBackColor = $True
	#
	# buttonRemoveIgnored
	#
	$buttonRemoveIgnored.Location = '342, 140'
	$buttonRemoveIgnored.Name = "buttonRemoveIgnored"
	$buttonRemoveIgnored.Size = '58, 23'
	$buttonRemoveIgnored.TabIndex = 4
	$buttonRemoveIgnored.Text = "Remove"
	$buttonRemoveIgnored.UseVisualStyleBackColor = $True
	$buttonRemoveIgnored.add_Click($buttonRemoveIgnored_Click)
	#
	# buttonAddIgnored
	#
	$buttonAddIgnored.Location = '194, 140'
	$buttonAddIgnored.Name = "buttonAddIgnored"
	$buttonAddIgnored.Size = '60, 23'
	$buttonAddIgnored.TabIndex = 3
	$buttonAddIgnored.Text = "Add"
	$buttonAddIgnored.UseVisualStyleBackColor = $True
	$buttonAddIgnored.add_Click($buttonAddIgnored_Click)
	#
	# txtPermIgnore
	#
	$txtPermIgnore.Location = '194, 114'
	$txtPermIgnore.Name = "txtPermIgnore"
	$txtPermIgnore.Size = '206, 20'
	$txtPermIgnore.TabIndex = 2
	$tooltipAll.SetToolTip($txtPermIgnore, "When reporting on permissions these users automatically get excluded from results.")
	#
	# labelIgnoredUsers
	#
	$labelIgnoredUsers.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelIgnoredUsers.Location = '194, 16'
	$labelIgnoredUsers.Name = "labelIgnoredUsers"
	$labelIgnoredUsers.Size = '117, 23'
	$labelIgnoredUsers.TabIndex = 8
	$labelIgnoredUsers.Text = "Ignored Users"
	#
	# labelIncludedReportElemen
	#
	$labelIncludedReportElemen.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelIncludedReportElemen.Location = '6, 16'
	$labelIncludedReportElemen.Name = "labelIncludedReportElemen"
	$labelIncludedReportElemen.Size = '162, 23'
	$labelIncludedReportElemen.TabIndex = 7
	$labelIncludedReportElemen.Text = "Included Report Elements"
	#
	# listPermReportIgnore
	#
	$listPermReportIgnore.FormattingEnabled = $True
	[void]$listPermReportIgnore.Items.Add("NT AUTHORITY\SYSTEM")
	[void]$listPermReportIgnore.Items.Add("NT AUTHORITY\SELF")
	$listPermReportIgnore.Location = '194, 42'
	$listPermReportIgnore.Name = "listPermReportIgnore"
	$listPermReportIgnore.Size = '206, 69'
	$listPermReportIgnore.TabIndex = 1
	$tooltipAll.SetToolTip($listPermReportIgnore, "Inherited or not these accounts will not be included on permissions reports")
	#
	# listPermReportOptions
	#
	$listPermReportOptions.FormattingEnabled = $True
	[void]$listPermReportOptions.Items.Add("Mailbox Summary Information")
	[void]$listPermReportOptions.Items.Add("Full Access Permissions")
	[void]$listPermReportOptions.Items.Add("Send On Behalf Permissions")
	[void]$listPermReportOptions.Items.Add("Send As Permissions")
	[void]$listPermReportOptions.Items.Add("Calendar Permissions")
	[void]$listPermReportOptions.Items.Add("Mailbox Rule - Forwarding")
	[void]$listPermReportOptions.Items.Add("Mailbox Rule - Redirecting")
	$listPermReportOptions.Location = '6, 42'
	$listPermReportOptions.Name = "listPermReportOptions"
	$listPermReportOptions.SelectionMode = 'MultiSimple'
	$listPermReportOptions.Size = '182, 121'
	$listPermReportOptions.TabIndex = 0
	$tooltipAll.SetToolTip($listPermReportOptions, "Select included reports")
	#
	# tabReportDeliveryOptions
	#
	$tabReportDeliveryOptions.Controls.Add($grpReportFormat)
	$tabReportDeliveryOptions.Controls.Add($groupbox1)
	$tabReportDeliveryOptions.BackColor = 'ControlLight'
	$tabReportDeliveryOptions.Location = '4, 22'
	$tabReportDeliveryOptions.Name = "tabReportDeliveryOptions"
	$tabReportDeliveryOptions.Padding = '3, 3, 3, 3'
	$tabReportDeliveryOptions.Size = '412, 404'
	$tabReportDeliveryOptions.TabIndex = 1
	$tabReportDeliveryOptions.Text = "Delivery Options"
	#
	# grpReportFormat
	#
	$grpReportFormat.Controls.Add($labelReportStyle)
	$grpReportFormat.Controls.Add($comboReportStyle)
	$grpReportFormat.Location = '8, 344'
	$grpReportFormat.Name = "grpReportFormat"
	$grpReportFormat.Size = '396, 54'
	$grpReportFormat.TabIndex = 75
	$grpReportFormat.TabStop = $False
	$grpReportFormat.Text = "Format"
	$grpReportFormat.Visible = $False
	#
	# labelReportStyle
	#
	$labelReportStyle.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelReportStyle.Location = '3, 19'
	$labelReportStyle.Name = "labelReportStyle"
	$labelReportStyle.Size = '123, 20'
	$labelReportStyle.TabIndex = 80
	$labelReportStyle.Text = "Report Style"
	$labelReportStyle.TextAlign = 'MiddleRight'
	#
	# comboReportStyle
	#
	$comboReportStyle.DropDownStyle = 'DropDownList'
	$comboReportStyle.FormattingEnabled = $True
	[void]$comboReportStyle.Items.Add("Default")
	$comboReportStyle.Location = '143, 19'
	$comboReportStyle.Name = "comboReportStyle"
	$comboReportStyle.Size = '244, 21'
	$comboReportStyle.TabIndex = 0
	#
	# groupbox1
	#
	$groupbox1.Controls.Add($txtEmailSubject)
	$groupbox1.Controls.Add($labelSubject)
	$groupbox1.Controls.Add($labelEmailReport)
	$groupbox1.Controls.Add($chkEmailReport)
	$groupbox1.Controls.Add($labelReportName)
	$groupbox1.Controls.Add($labelSaveReport)
	$groupbox1.Controls.Add($txtReportName)
	$groupbox1.Controls.Add($labelReportFolder)
	$groupbox1.Controls.Add($chkSaveLocally)
	$groupbox1.Controls.Add($txtReportFolder)
	$groupbox1.Controls.Add($txtSMTPServer)
	$groupbox1.Controls.Add($buttonBrowseFolder)
	$groupbox1.Controls.Add($labelSMTPRelayServer)
	$groupbox1.Controls.Add($labelEmailSender)
	$groupbox1.Controls.Add($labelEmailRecipient)
	$groupbox1.Controls.Add($txtEmailRecipient)
	$groupbox1.Controls.Add($txtEmailSender)
	$groupbox1.BackColor = 'ControlLight'
	$groupbox1.FlatStyle = 'System'
	$groupbox1.Location = '9, 6'
	$groupbox1.Name = "groupbox1"
	$groupbox1.RightToLeft = 'No'
	$groupbox1.Size = '395, 247'
	$groupbox1.TabIndex = 74
	$groupbox1.TabStop = $False
	$groupbox1.Text = "Delivery"
	#
	# txtEmailSubject
	#
	$txtEmailSubject.Enabled = $False
	$txtEmailSubject.Location = '142, 61'
	$txtEmailSubject.Name = "txtEmailSubject"
	$txtEmailSubject.Size = '244, 20'
	$txtEmailSubject.TabIndex = 86
	#
	# labelSubject
	#
	$labelSubject.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelSubject.Location = '2, 60'
	$labelSubject.Name = "labelSubject"
	$labelSubject.Size = '123, 20'
	$labelSubject.TabIndex = 87
	$labelSubject.Text = "Subject"
	$labelSubject.TextAlign = 'MiddleRight'
	#
	# labelEmailReport
	#
	$labelEmailReport.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelEmailReport.Location = '35, 35'
	$labelEmailReport.Name = "labelEmailReport"
	$labelEmailReport.Size = '89, 20'
	$labelEmailReport.TabIndex = 85
	$labelEmailReport.Text = "Email Report"
	$labelEmailReport.TextAlign = 'MiddleRight'
	#
	# chkEmailReport
	#
	$chkEmailReport.Location = '142, 38'
	$chkEmailReport.Name = "chkEmailReport"
	$chkEmailReport.Size = '14, 17'
	$chkEmailReport.TabIndex = 84
	$chkEmailReport.UseVisualStyleBackColor = $True
	$chkEmailReport.add_CheckedChanged($chkEmailReport_CheckedChanged)
	#
	# labelReportName
	#
	$labelReportName.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelReportName.Location = '35, 187'
	$labelReportName.Name = "labelReportName"
	$labelReportName.Size = '89, 20'
	$labelReportName.TabIndex = 75
	$labelReportName.Text = "Report Name"
	$labelReportName.TextAlign = 'MiddleRight'
	#
	# labelSaveReport
	#
	$labelSaveReport.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelSaveReport.Location = '-3, 161'
	$labelSaveReport.Name = "labelSaveReport"
	$labelSaveReport.Size = '128, 20'
	$labelSaveReport.TabIndex = 82
	$labelSaveReport.Text = "Save Report"
	$labelSaveReport.TextAlign = 'MiddleRight'
	#
	# txtReportName
	#
	$txtReportName.Enabled = $False
	$txtReportName.Location = '142, 187'
	$txtReportName.Name = "txtReportName"
	$txtReportName.Size = '244, 20'
	$txtReportName.TabIndex = 83
	$txtReportName.Text = "Report.html"
	#
	# labelReportFolder
	#
	$labelReportFolder.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelReportFolder.Location = '35, 211'
	$labelReportFolder.Name = "labelReportFolder"
	$labelReportFolder.Size = '89, 20'
	$labelReportFolder.TabIndex = 80
	$labelReportFolder.Text = "Report Folder"
	$labelReportFolder.TextAlign = 'MiddleRight'
	#
	# chkSaveLocally
	#
	$chkSaveLocally.Location = '142, 164'
	$chkSaveLocally.Name = "chkSaveLocally"
	$chkSaveLocally.Size = '14, 17'
	$chkSaveLocally.TabIndex = 81
	$chkSaveLocally.UseVisualStyleBackColor = $True
	$chkSaveLocally.add_CheckedChanged($chkSaveLocally_CheckedChanged)
	#
	# txtReportFolder
	#
	$txtReportFolder.AutoCompleteMode = 'SuggestAppend'
	$txtReportFolder.AutoCompleteSource = 'FileSystemDirectories'
	$txtReportFolder.Enabled = $False
	$txtReportFolder.Location = '142, 213'
	$txtReportFolder.Name = "txtReportFolder"
	$txtReportFolder.Size = '207, 20'
	$txtReportFolder.TabIndex = 72
	$txtReportFolder.Text = "."
	#
	# txtSMTPServer
	#
	$txtSMTPServer.Enabled = $False
	$txtSMTPServer.Location = '142, 86'
	$txtSMTPServer.Name = "txtSMTPServer"
	$txtSMTPServer.Size = '244, 20'
	$txtSMTPServer.TabIndex = 74
	#
	# buttonBrowseFolder
	#
	$buttonBrowseFolder.Location = '356, 211'
	$buttonBrowseFolder.Name = "buttonBrowseFolder"
	$buttonBrowseFolder.Size = '30, 23'
	$buttonBrowseFolder.TabIndex = 73
	$buttonBrowseFolder.Text = "..."
	$buttonBrowseFolder.UseVisualStyleBackColor = $True
	$buttonBrowseFolder.add_Click($buttonBrowseFolder_Click)
	#
	# labelSMTPRelayServer
	#
	$labelSMTPRelayServer.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelSMTPRelayServer.Location = '2, 85'
	$labelSMTPRelayServer.Name = "labelSMTPRelayServer"
	$labelSMTPRelayServer.Size = '123, 20'
	$labelSMTPRelayServer.TabIndex = 75
	$labelSMTPRelayServer.Text = "SMTP Relay Server"
	$labelSMTPRelayServer.TextAlign = 'MiddleRight'
	#
	# labelEmailSender
	#
	$labelEmailSender.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelEmailSender.Location = '35, 137'
	$labelEmailSender.Name = "labelEmailSender"
	$labelEmailSender.Size = '90, 20'
	$labelEmailSender.TabIndex = 78
	$labelEmailSender.Text = "Email Sender"
	$labelEmailSender.TextAlign = 'MiddleRight'
	#
	# labelEmailRecipient
	#
	$labelEmailRecipient.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelEmailRecipient.Location = '2, 111'
	$labelEmailRecipient.Name = "labelEmailRecipient"
	$labelEmailRecipient.Size = '123, 20'
	$labelEmailRecipient.TabIndex = 79
	$labelEmailRecipient.Text = "Email Recipient"
	$labelEmailRecipient.TextAlign = 'MiddleRight'
	#
	# txtEmailRecipient
	#
	$txtEmailRecipient.Enabled = $False
	$txtEmailRecipient.Location = '142, 112'
	$txtEmailRecipient.Name = "txtEmailRecipient"
	$txtEmailRecipient.Size = '244, 20'
	$txtEmailRecipient.TabIndex = 76
	#
	# txtEmailSender
	#
	$txtEmailSender.Enabled = $False
	$txtEmailSender.Location = '142, 138'
	$txtEmailSender.Name = "txtEmailSender"
	$txtEmailSender.Size = '244, 20'
	$txtEmailSender.TabIndex = 77
	#
	# tabAbout
	#
	$tabAbout.Controls.Add($richtextboxAbout1)
	$tabAbout.BackColor = 'ControlLight'
	$tabAbout.Location = '4, 22'
	$tabAbout.Name = "tabAbout"
	$tabAbout.Size = '412, 404'
	$tabAbout.TabIndex = 4
	$tabAbout.Text = "About"
	#
	# richtextboxAbout1
	#
	$richtextboxAbout1.Location = '3, 4'
	$richtextboxAbout1.Name = "richtextboxAbout1"
	$richtextboxAbout1.ScrollBars = 'Vertical'
	$richtextboxAbout1.Size = '406, 397'
	$richtextboxAbout1.TabIndex = 0
	$richtextboxAbout1.Text = ""
	#
	# timerFadeIn
	#
	$timerFadeIn.add_Tick($timerFadeIn_Tick)
	#
	# folderbrowserdialog1
	#
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
	# Created by:   Zachary Loeber
	# Filename: Globals.ps1
	# Description: These are used across both the gui and the called script
	#              for storing and loading script state data.
	#========================================================================
	
	# RequiredSnapins
	$Snapins=@(’Microsoft.Exchange.Management.PowerShell.E2010’)
	
	#Our base variables
	$varEmailReport=$false
	$varEmailSubject=""
	$varEmailRecipient=""
	$varEmailSender=""
	$varSMTPServer=""
	$varSaveReportsLocally=$true
	$varReportName="report.html"
	$varReportFolder="."
	$varServer="localhost"
	$varUseCurrentUser=$true
	$varUser=""
	$varPassword=""
	$varScopeEnterprise=$true
	$varScopeDAG=""
	$varScopeServer=""
	$varScopeMailboxDatabase=""
	
	#Different Report Option defaults
	$varMailboxReportPermissions = @()
	$hashAllPermReports = @(@{Option="Mailbox Summary Information"; Selected=$true},
	                        @{Option="Full Access Permissions"; Selected=$true},
	                        @{Option="Send On Behalf Permissions"; Selected=$true},
	                        @{Option="Send As Permissions"; Selected=$true},
							@{Option="Calendar Permissions"; Selected=$true},
	                        @{Option="Mailbox Rule - Forwarding"; Selected=$true},
	                        @{Option="Mailbox Rule - Redirecting"; Selected=$true})
	        
	foreach ($hashOption in $hashAllPermReports)
	{
	    $Newobject = New-Object PSObject -Property $hashOption
	    $varMailboxReportPermissions += $Newobject
	}
	
	$varMailboxReportIgnoredUsers = @(  "NT AUTHORITY\SYSTEM",
	                                    "NT AUTHORITY\SELF")
	$varIncludeInherited = $true
	$varExcludeZeroResults = $false
	$varExcludeUnknownUsers = $true
	$varMailboxRuleForwarding = $true
	$varMailboxRuleRedirecting = $true
	$varSummaryReport = $true
	$varFullAccessReport = $true
	$varSendAsReport = $true
	$varSendOnBehalfReport = $true
	$varCalendarPermReport = $true
	$varFlagWarnings = $true
	$varMailboxSizeWarning = 512
	$varMailboxSizeAlert = 1024
	$varDeletedSizeWarning = 512
	$varDeletedSizeAlert = 1024
	
	
	#For each variable an xml attribute should exist
	$ConfigTemplate = 
	@"
<Configuration>
    <EmailReport>{0}</EmailReport>
    <EmailSubject>{1}</EmailSubject>
	<EmailRecipient>{2}</EmailRecipient>
	<EmailSender>{3}</EmailSender>
	<SMTPServer>{4}</SMTPServer>
	<SaveReportsLocally>{5}</SaveReportsLocally>
	<ReportName>{6}</ReportName>
	<ReportFolder>{7}</ReportFolder>
    <Server>{8}</Server>
    <UseCurrentUser>{9}</UseCurrentUser>
    <User>{10}</User>
    <Password>{11}</Password>
    <ScopeEnterprise>{12}</ScopeEnterprise>
    <ScopeDAG>{13}</ScopeDAG>
    <ScopeServer>{14}</ScopeServer>
    <ScopeMailboxDatabase>{15}</ScopeMailboxDatabase>
    <IncludeInherited>{16}</IncludeInherited>
    <ExcludeZeroResults>{17}</ExcludeZeroResults>
    <ExcludeUnknownUsers>{18}</ExcludeUnknownUsers>
	<MailboxRuleForwarding>{19}</MailboxRuleForwarding>
	<MailboxRuleRedirecting>{20}</MailboxRuleRedirecting>
	<SummaryReport>{21}</SummaryReport>
	<FullAccessReport>{22}</FullAccessReport>
	<SendAsReport>{23}</SendAsReport>
	<SendOnBehalfReport>{24}</SendOnBehalfReport>
	<CalendarPermReport>{25}</CalendarPermReport>
	<FlagWarnings>{26}</FlagWarnings>
	<MailboxSizeWarning>{27}</MailboxSizeWarning>
	<MailboxSizeAlert>{28}</MailboxSizeAlert>
	<DeletedSizeWarning>{29}</DeletedSizeWarning>
	<DeletedSizeAlert>{30}</DeletedSizeAlert>
</Configuration>
"@
	
	# Exchange specific globals that do not get saved
	$EXConnected = $false
	#$MailboxServerScope = @()
	$datagridMailboxes = @()
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
	
	#Config file location
	$ConfigFile = $ScriptDirectory + "\Config.xml"
	
	Function Get-SaveData
	{
	    [xml]$x=($ConfigTemplate) -f
	        $varEmailReport,`
	        $varEmailSubject,`
	        $varEmailRecipient,`
	        $varEmailSender,`
	        $varSMTPServer,`
	        $varSaveReportsLocally,`
	        $varReportName,`
	        $varReportFolder,`
	        $varServer,`
	        $varUseCurrentUser,`
	        $varUser,`
	        $varPassword,`
	        $varScopeEnterprise,`
	        $varScopeDAG,`
	        $varScopeServer,`
	        $varScopeMailboxDatabase,`
			$varIncludeInherited,`
			$varExcludeZeroResults,`
			$varExcludeUnknownUsers,`
			$varMailboxRuleForwarding, `
			$varMailboxRuleRedirecting, `
			$varSummaryReport, `
			$varFullAccessReport, `
			$varSendAsReport, `
			$varSendOnBehalfReport, `
			$varCalendarPermReport, `
			$varFlagWarnings, `
			$varMailboxSizeWarning, `
			$varMailboxSizeAlert, `
			$varDeletedSizeWarning, `
			$varDeletedSizeAlert
	    return $x
	}
	
	function Load-Config
	{
		if (Test-Path $ConfigFile)
		{
			[xml]$configuration = get-content $($ConfigFile)
	        $Script:varEmailReport = [System.Convert]::ToBoolean($configuration.Configuration.EmailReport)
	        $Script:varEmailSubject = $configuration.Configuration.EmailSubject
	        $Script:varEmailRecipient = $configuration.Configuration.EmailRecipient
	        $Script:varEmailSender = $configuration.Configuration.EmailSender
	        $Script:varSMTPServer = $configuration.Configuration.SMTPServer
	        $Script:varSaveReportsLocally = [System.Convert]::ToBoolean($configuration.Configuration.SaveReportsLocally)
	        $Script:varReportName = $configuration.Configuration.ReportName
	        $Script:varReportFolder = $configuration.Configuration.ReportFolder
	        $Script:varServer = $configuration.Configuration.Server
	        $Script:varUseCurrentUser = [System.Convert]::ToBoolean($configuration.Configuration.UseCurrentUser)
	        $Script:varUser = $configuration.Configuration.User
	        $Script:varScopeEnterprise = [System.Convert]::ToBoolean($configuration.Configuration.ScopeEnterprise)
	        $Script:varScopeDAG = $configuration.Configuration.ScopeDAG
	        $Script:varScopeServer = $configuration.Configuration.ScopeServer
	        $Script:varScopeMailboxDatabase = $configuration.Configuration.ScopeMailboxDatabase
			$Script:varIncludeInherited = [System.Convert]::ToBoolean($configuration.Configuration.IncludeInherited)
			$Script:varExcludeZeroResults = [System.Convert]::ToBoolean($configuration.Configuration.ExcludeZeroResults)
			$Script:varExcludeUnknownUsers = [System.Convert]::ToBoolean($configuration.Configuration.ExcludeUnknownUsers)
			$Script:varMailboxRuleForwarding = [System.Convert]::ToBoolean($configuration.Configuration.MailboxRuleForwarding)
			$Script:varMailboxRuleRedirecting = [System.Convert]::ToBoolean($configuration.Configuration.MailboxRuleRedirecting)
			$Script:varSummaryReport = [System.Convert]::ToBoolean($configuration.Configuration.SummaryReport)
			$Script:varFullAccessReport = [System.Convert]::ToBoolean($configuration.Configuration.FullAccessReport)
			$Script:varSendAsReport = [System.Convert]::ToBoolean($configuration.Configuration.SendAsReport)
			$Script:varSendOnBehalfReport = [System.Convert]::ToBoolean($configuration.Configuration.SendOnBehalfReport)
			$Script:varCalendarPermReport = [System.Convert]::ToBoolean($configuration.Configuration.CalendarPermReport)
			$Script:varFlagWarnings = [System.Convert]::ToBoolean($configuration.Configuration.FlagWarnings)
			$Script:varMailboxSizeWarning = $configuration.Configuration.MailboxSizeWarning
			$Script:varMailboxSizeAlert = $configuration.Configuration.MailboxSizeAlert
			$Script:varDeletedSizeWarning = $configuration.Configuration.DeletedSizeWarning
			$Script:varDeletedSizeAlert = $configuration.Configuration.DeletedSizeAlert
	        $Script:varMailboxReportPermissions = @()
	        foreach ($mboxoption in $configuration.Configuration.MailboxReportPermissions)
	        {
	            $hash = @{
					Option = $mboxoption.Option;
					Selected = [System.Convert]::ToBoolean($mboxoption.Selected)
				}			
	            $Newobject = New-Object PSObject -Property $hash
	            $Script:varMailboxReportPermissions += $Newobject
			}
	        $Script:varMailboxReportIgnoredUsers = @()
	        foreach ($ignoreduser in $configuration.Configuration.MailboxReportIgnoredUser)
	        {
	            $Script:varMailboxReportIgnoredUsers += $ignoreduser.User
			}
	        
	        Return $true
		}
	    else
	    {
	        Return $false
	    }
	}
	
	# Save exceptions
	function Save-Config
	{
	    $SanitizedConfig = $true
		if (($varEmailReport) -and`
	        (($varEmailSubject -eq "") -or`
			 ($varEmailRecipient -eq "") -or`
	         ($varEmailSender -eq "") -or`
	         ($varSMTPServer -eq "")))
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("You selected to send an email but didn't fill in the right stuff to make it happen buddy.","Sorry, try again.")
	        $SanitizedConfig = $false
		}
		elseif (($varSaveReportsLocally) -and`
				(($varReportName -eq "") -or`
				 ($varReportFolder -eq "")))
		{
			#[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
			[void][System.Windows.Forms.MessageBox]::Show("You selected to not save locally (so are assumed to be attempting to email the reports) but didn't fill in email configuration information.","Sorry, not going to do it.")
	        $SanitizedConfig = $false
		}
	 
	    if ($SanitizedConfig)
		{
			# save the data
			[xml]$x=Get-SaveData
	        foreach ($mboxperm in $varMailboxReportPermissions)
	        {
	            $newpermoption = $x.CreateElement("MailboxReportPermissions")
	            $newpermoption.InnerXML = "<Option>$($mboxperm.Option)</Option><Selected>$($mboxperm.Selected)</Selected>"
	            $x.Configuration.AppendChild($newpermoption)
			}
	        foreach($ignoreduser in $varMailboxReportIgnoredUsers)
	        {
				$ignoreduser2 = $x.CreateElement("MailboxReportIgnoredUser")
	            $ignoreduser2.InnerXML = "<User>$($ignoreduser)</User>"
	            $x.Configuration.AppendChild($ignoreduser2)
			}
	        $x.save($ConfigFile)
	        Return $true
		}
	    else
	    {
	        Return $false
		}
	}
	
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
	
	Function Get-MailboxList {
	    [CmdletBinding()] 
	    param ( 
	        [Parameter(HelpMessage="Return all mailboxes in a specific DAG.")]
	        [String]$DAGName = '',
	        [Parameter(HelpMessage="Return all mailboxes on a specific server.")]
	        [String]$ServerName = '', 
	        [Parameter(HelpMessage="Return all mailboxes in a specific mailbox.")]
	        [String]$DatabaseName = '',
	        [Parameter(HelpMessage="Return all mailboxes in the enterprise.")]
	        [switch]$WholeEnterprise,
	        [Parameter(HelpMessage="Return specific mailboxes.")]
	        [string[]]$Mailboxes,
	        [Parameter(HelpMessage="Returns only the mailbox identities instead of an array of mailbox objects.")]
	        [switch]$ReturnIdentitiesOnly
	    )
	    BEGIN {
	        $date = get-date -Format MM-dd-yyyy
	    }
	    PROCESS {
	        $Results = @()
	        if ($Mailboxes) {
	            $MailboxInput = @()
	            $MailboxInput += $Mailboxes
	        }
	
	        if ($Mailboxes.Count -ge 1) {
	            Foreach ($Mbox in $MailboxInput) {
	                try {
	                    $_Mbox = Get-Mailbox $Mbox -ErrorAction 'Stop'
	                    $Results += $_Mbox
	                }
	                catch {
	                    $time = get-date -Format hh.mm
	                    $erroroutput = "$date;$time;$Mbox;$_"
	                    Write-Warning $erroroutput
	                }
	            }
	        }
	        elseif ($DatabaseName -ne '') {
	            try {
	                $Results = @(Get-Mailbox -Database $DatabaseName -ResultSize Unlimited)
	            }
	            catch {
	                $time = get-date -Format hh.mm
	                $erroroutput = "$date;$time;$DatabaseName;$_"
	                Write-Warning $erroroutput
	            }
	                  
	        }
	        elseif ($ServerName -ne '') {
	            try {
	                $Results += @(Get-Mailbox -Server $ServerName -ResultSize Unlimited)
	            }
	            catch {
	                $time = get-date -Format hh.mm
	                $erroroutput = "$date;$time;$ServerName;$_"
	                Write-Warning $erroroutput
	            }
	        }
	        elseif ($DAGName -ne '') {
	            try {
	                $Servers = @(Get-DatabaseAvailabilityGroup $DAGName | foreach {$_.Servers})
	                foreach ($Server in $Servers) {
	                    $Results += @(Get-Mailbox -Server $Server -ResultSize Unlimited)
	                }
	            }
	            catch {
	                $time = get-date -Format hh.mm
	                $erroroutput = "$date;$time;$DAGName;$_"
	                Write-Warning $erroroutput
	            }
	        }
	        elseif ($WholeEnterprise) {
	            $Results += @(Get-Mailbox -ResultSize Unlimited)
	        }
	        
	        if ($ReturnIdentitiesOnly) {
	            $Results = @($Results | %{[string]$_.Identity})
	        }
	        Return $Results 
	    }
	}
	
	function Colorize-Table {
	<# 
	.SYNOPSIS 
	Colorize-Table 
	 
	.DESCRIPTION 
	Create an html table and colorize individual cells or rows of an array of objects based on row header and value. Optionally, you can also
	modify an existing html document or change only the styles of even or odd rows.
	 
	.PARAMETER  InputObject 
	An array of objects (ie. (Get-process | select Name,Company) 
	 
	.PARAMETER  Column 
	The column you want to modify. (Note: If the parameter ColorizeMethod is not set to ByValue the 
	Column parameter is ignored)
	
	.PARAMETER ScriptBlock
	Used to perform custom cell evaluations such as -gt -lt or anything else you need to check for in a
	table cell element. The scriptblock must return either $true or $false and is, by default, just
	a basic -eq comparisson. You must use the variables as they are used in the following example.
	(Note: If the parameter ColorizeMethod is not set to ByValue the ScriptBlock parameter is ignored)
	
	[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
	
	$args[0] will be the cell value in the table
	$args[1] will be the value to compare it to
	
	Strong typesetting is encouraged for accuracy.
	
	.PARAMETER  ColumnValue 
	The column value you will modify if ScriptBlock returns a true result. (Note: If the parameter 
	ColorizeMethod is not set to ByValue the ColumnValue parameter is ignored)
	 
	.PARAMETER  Attr 
	The attribute to change should ColumnValue be found in the Column specified. 
	- A good example is using "style" 
	 
	.PARAMETER  AttrValue 
	The attribute value to set when the ColumnValue is found in the Column specified 
	- A good example is using "background: red;" 
	 
	.EXAMPLE 
	This will highlight the process name of Dropbox with a red background. 
	
	$TableStyle = @'
	<title>Process Report</title> 
	    <style>             
	    BODY{font-family: Arial; font-size: 8pt;} 
	    H1{font-size: 16px;} 
	    H2{font-size: 14px;} 
	    H3{font-size: 12px;} 
	    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;} 
	    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;} 
	    TD{border: 1px solid black; padding: 5px;} 
	    </style>
	'@
	
	$tabletocolorize = $(Get-Process | ConvertTo-Html -Head $TableStyle) 
	$colorizedtable = Colorize-Table $tabletocolorize -Column "Name" -ColumnValue "Dropbox" -Attr "style" -AttrValue "background: red;"
	$colorizedtable | Out-File "$pwd/testreport.html" 
	ii "$pwd/testreport.html"
	
	You can also strip out just the table at the end if you are working with multiple tables in your report:
	if ($colorizedtable -match '(?s)<table>(.*)</table>')
	{
	    $result = $matches[0]
	}
	
	.EXAMPLE 
	Using the same $TableStyle variable above this will create a table of top 5 processes by memory usage,
	color the background of a whole row yellow for any process using over 150Mb and red if over 400Mb.
	
	$tabletocolorize = $(get-process | select -Property ProcessName,Company,@{Name="Memory";Expression={[math]::truncate($_.WS/ 1Mb)}} | Sort-Object Memory -Descending | Select -First 5 ) 
	
	[scriptblock]$scriptblock = {[int]$args[0] -gt [int]$args[1]}
	$testreport = Colorize-Table $tabletocolorize -Column "Memory" -ColumnValue 150 -Attr "style" -AttrValue "background:yellow;" -ScriptBlock $ScriptBlock -HTMLHead $TableStyle -WholeRow $true
	$testreport = Colorize-Table $testreport -Column "Memory" -ColumnValue 400 -Attr "style" -AttrValue "background:red;" -ScriptBlock $ScriptBlock -WholeRow $true
	$testreport | Out-File "$pwd/testreport.html" 
	ii "$pwd/testreport.html"
	
	.NOTES 
	If you are going to convert something to html with convertto-html in powershell v2 there is a bug where the  
	header will show up as an asterick if you only are converting one object property. 
	
	This script is a modification of something I found by some rockstar named Jaykul at this site
	http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as
	
	I believe that .Net 4.0 is a requirement for using the Linq libraries
	
	.LINK 
	http://www.the-little-things.net 
	#> 
	[CmdletBinding(DefaultParameterSetName = "ObjectSet")] 
	param ( 
	    [Parameter( Position=0,
	                Mandatory=$true, 
	                ValueFromPipeline=$true, 
	                ParameterSetName="ObjectSet")]
	        [PSObject[]]$InputObject, 
	    [Parameter( Position=0, 
	                Mandatory=$true, 
	                ValueFromPipeline=$true, 
	                ParameterSetName="StringSet")]
	        [String[]]$InputString='', 
	    [Parameter( Position=1 )]
	        [String]$Column="Name", 
	    [Parameter( Position=2 )] 
	        $ColumnValue=0,
	    [Parameter( Position=3 )]
	        [ScriptBlock]$ScriptBlock = {[string]$args[0] -eq [string]$args[1]}, 
	    [Parameter( Position=4, 
	                Mandatory=$true )]
	        [String]$Attr, 
	    [Parameter( Position=5, 
	                Mandatory=$true )]
	        [String]$AttrValue, 
	    [Parameter( Position=6 )]
	        [Bool]$WholeRow=$false, 
	    [Parameter( Position=7, 
	                ParameterSetName="ObjectSet")] 
	        [String]$HTMLHead='<title>HTML Table</title>',
	    [Parameter( Position=8 )]
	    [ValidateSet('ByValue','ByEvenRows','ByOddRows')]
	        [String]$ColorizeMethod='ByValue'
	    )
	    
	BEGIN 
	{ 
	    # A little note on Add-Type, this adds in the assemblies for linq with some custom code. The first time this 
	    # is run in your powershell session it is compiled and loaded into your session. If you run it again in the same
	    # session and the code was not changed at all powershell skips the command (otherwise recompiling code each time
	    # the function is called in a session would be pretty ineffective so this is by design). If you make any changes
	    # to the code, even changing one space or tab, it is detected as new code and will try to reload the same namespace
	    # which is not allowed and will cause an error. So if you are debugging this or changing it up, either change the
	    # namespace as well or exit and restart your powershell session.
	    #
	    # And some notes on the actual code. It is my first jump into linq (or C# for that matter) so if it looks not so 
	    # elegant or there is a better way to do this I'm all ears. I define four methods which names are self-explanitory:
	    # - GetElementByIndex
	    # - GetElementByValue
	    # - GetOddElements
	    # - GetEvenElements
	    $LinqCode = @"
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByIndex(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, int index)
    {
        return doc.Descendants(element)
                .Where  (e => e.NodesBeforeSelf().Count() == index)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetElementByValue(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element, string value)
    {
        return  doc.Descendants(element) 
                .Where  (e => e.Value == value)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetOddElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 != 0)
                .Select (e => e);
    }
    public static System.Collections.Generic.IEnumerable<System.Xml.Linq.XElement> GetEvenElements(System.Xml.Linq.XContainer doc, System.Xml.Linq.XName element)
    {
        return doc.Descendants(element)
                .Where  ((e,i) => i % 2 == 0)
                .Select (e => e);
    }
"@
	
	    Add-Type -ErrorAction SilentlyContinue -Language CSharpVersion3 `
	    -ReferencedAssemblies System.Xml, System.Xml.Linq `
	    -UsingNamespace System.Linq `
	    -Name XUtilities `
	    -Namespace Huddled `
	    -MemberDefinition $LinqCode
	    
	    $Objects = @() 
	} 
	 
	PROCESS 
	{ 
	    $Objects += $InputObject 
	} 
	 
	END 
	{ 
	    # Convert our data to x(ht)ml 
	    if ($InputString)    # If a string was passed just parse it 
	    { 
	        $xml = [System.Xml.Linq.XDocument]::Parse("$InputString")  
	    } 
	    else    # Otherwise we have to convert it to html first 
	    { 
	        $xml = [System.Xml.Linq.XDocument]::Parse("$($Objects | ConvertTo-Html -Head $HTMLHead)")
	    } 
	    
	    switch ($ColorizeMethod) {
	        "ByEvenRows" {
	            $evenrows = [Huddled.XUtilities]::GetEvenElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
	            foreach ($row in $evenrows)
	            {
	                $row.SetAttributeValue($Attr, $AttrValue)
	            }            
	        }
	
	        "ByOddRows" {
	            $oddrows = [Huddled.XUtilities]::GetOddElements($xml, "{http://www.w3.org/1999/xhtml}tr")    
	            foreach ($row in $oddrows)
	            {
	                $row.SetAttributeValue($Attr, $AttrValue)
	            }
	        }
	        "ByValue" {
	            # Find the index of the column you want to format 
	            $ColumnLoc = [Huddled.XUtilities]::GetElementByValue($xml, "{http://www.w3.org/1999/xhtml}th",$Column) 
	            $ColumnIndex = $ColumnLoc | Foreach-Object{($_.NodesBeforeSelf() | Measure-Object).Count} 
	    
	            # Process each xml element based on the index for the column we are highlighting 
	            switch([Huddled.XUtilities]::GetElementByIndex($xml, "{http://www.w3.org/1999/xhtml}td", $ColumnIndex)) 
	            { 
	                {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $ColumnValue))} {
	                    if ($WholeRow)
	                    {
	                        $_.Parent.SetAttributeValue($Attr, $AttrValue)
	                    }
	                    else
	                    {
	                        $_.SetAttributeValue($Attr, $AttrValue)
	                    }
	                }
	            }
	        }
	    }
	    Return $xml.Document.ToString()
	}
	}
	
	Function Get-CalendarPermission
	{
	<#
	.Synopsis
	    Retrieves a list of mailbox calendar permissions
	.DESCRIPTION
	    Get-CalendarPermission uses the exchange 2010 snappin to get a list of permissions for mailboxes in an exchange environment.
	    As different languages spell calendar differently this script first pulls the actual name of the calendar by using
	    get-mailboxfolderstatistics and has proven to work across multi-lingual organizations.
	.PARAMETER Mailbox
	    One or more mailbox names.
	.PARAMETER LogErrors
	    By default errors are not logged. Use -LogErrors to enable logging errors.
	.PARAMETER ErrorLog
	    When used with -LogErrors it specifies the full path and location for the ErrorLog. Defaults to "D:\errorlog.txt"
	.LINK
	    
	.NOTES        
	Name        :   Get Exchange Calendar Permissions
	Last edit   :   April 14th 2013
	Version     :   1.2.0 May 6 2013    :   Fixed issue where a mailbox name produces more than one mailbox
	                1.1.0 April 24 2013 :   Used new script template from http://blog.bjornhouben.com
	                1.0.0 March 10 2013 :   Created script
	
	Author      :   Zachary Loeber
	Website     :   http://www.the-little-things.net
	Linkedin    :   http://nl.linkedin.com/in/zloeber
	Keywords    :   Exchange, Calendar, Permissions Report
	Disclaimer  :   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
	                any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
	                performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
	                (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
	                or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
	
	Copyright   :   I believe in sharing knowledge, so this script and its use is subject to : http://creativecommons.org/licenses/by-sa/3.0/
	
	.EXAMPLE
	    Get-CalendarPermission -MailboxName "Test User1" -LogErrors -logfile "C:\logfile.txt" -Verbose
	    
	    Description
	    -----------
	    Gets the calendar permissions for "Test User1", logs errors to "C:\myerrorlog.txt" and shows verbose information.
	    
	.EXAMPLE
	    Get-CalendarPermission -MailboxName "user1","user2" -LogErrors -ErrorLog "C:\myerrorlog.txt" | Format-List
	    
	    Description
	    -----------
	    
	    Gets the calendar permissions for "user1" and "user2", logs errors to "C:\myerrorlog.txt" and returns the info as a format-list.
	
	.EXAMPLE
	    (Get-Mailbox -Database "MDB1") | Get-CalendarPermission -LogErrors -Logfile "C:\myerrorlog.txt" | Format-Table Mailbox,User,Permission
	    
	    Description
	    -----------
	    Gets all mailboxes in the MDB1 database and pipes it to Get-CalendarPermission. Get-CalendarPermission logs errors to "C:\myerrorlog.txt" and returns the info as an autosized format-table containing the Mailbox,User, and Permission
	#>
	    [CmdletBinding()]
	    param(
	        [Parameter( Mandatory=$True,
	                    ValueFromPipeline=$True,
	                    Position=0,
	                    HelpMessage="Enter an Exchange mailbox name")]
	        [string[]]$MailboxName,
	        [Parameter( Position=1,
	                    HelpMessage='Enter the full path for your log file. By example: "C:\Windows\log.txt"')]
	        [Alias("LogFile")]
	            [String]$ErrorLog = ".\errorlog.txt",    
	        [switch]$LogErrors
	    )
	    PROCESS
	    {
	        $Mboxes = @()
	        $Mboxes += $MailboxName
	        Foreach($Mailbox in $Mboxes)
	        {
	            TRY
	            { 
	                $Mbox = @(Get-Mailbox $Mailbox -erroraction Stop)
	                $CheckSuccesful = $True
	            }
	            CATCH
	            {
	                $CheckSuccesful = $False
	                $date = get-date -Format dd-MM-yyyy
	                $time = get-date -Format hh.mm
	                $erroroutput = "$date;$time;$Mailbox;MailboxError;$_.Exception.Message"
	
	                Write-Warning $erroroutput
	
	                IF($LogErrors -eq $True)
	                {
	                    Write-Verbose "Writing error for Mailbox : $Mailbox to $ErrorLog"
	                    $ErrorOutput | Out-File $ErrorLog -Append -Encoding ASCII
	                }
	            }
	             
	            IF ($CheckSuccesful -eq $True) #If Mailbox was found keep processing
	            {
	                ForEach ($MailUser in $Mbox) {
	                    # Construct the full path to the calendar folder regardless of the language
	                    $Calfolder = $MailUser.Name
	                    $Calfolder = $Calfolder + ':\'
	                    $CalFolder = $Calfolder + [string](Get-MailboxFolderStatistics $MailUser.Identity -folderscope calendar).Name
	                    $CalPerm = Get-MailboxFolderPermission $Calfolder
	                    $Results = @()
	                    foreach ($Perm in $CalPerm)
	                    {
	                        $TempHash = @{
	                            'Mailbox'=$MailUser.Name;
	                            'User'=$Perm.User;
	                            'Permission'=$Perm.AccessRights;
	                        }
	                        $Tempobject = New-Object PSObject -Property $TempHash
	                        $Results = $Results + $Tempobject
	                    }
	                }
	                $Results
	            }
	        }
	    }
	}
#endregion Source: Globals.ps1

#Start the application
Main ($CommandLine)
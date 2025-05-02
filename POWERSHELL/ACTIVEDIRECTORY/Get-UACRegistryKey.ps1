#\/******/\********\/********/\********\/********/\********\/********/\********\/******\/
#/\()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()/\
#\/()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()\/
#/\******\/********/\********\/********/\********\/********/\********\/********/\******/\
#\-\
#-\-Author:  Chris Rakowitz
#\-\Purpose:  Extract the Registry information for the User Account Control settings.
#/-/Date:  January 11, 2017
#-/-Version:  1.00
#/-/
#\/******/\********\/********/\********\/********/\********\/********/\********\/******\/
#/\()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()/\
#\/()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()()/\()()()()\/()()()\/
#/\******\/********/\********\/********/\********\/********/\********\/********/\******/\

<#
	.SYNOPSIS
	Extracts the current Registry Values for User Account Control.
	
	.DESCRIPTION
	The User Account Control (UAC) is used for controlling how Administrator and Standard
	User accounts react to attempts to install software.  This includes settings like
	dimming the screen when asking for credentials and Admin Approval Mode which requires
	even Administrators to click Yes/No when installing software.  These settings are
	helpful for preventing Malware from installing software un-noticed.
	
	This script reads the registry to extract the specific settings that are curently set
	on a Windows Vista or later PC.  It also extracts some other settings if they exist,
	such as the Ctrl+Alt+Del screen, Not showing the last user to log in to the computer, 
	whether a user can undock or shutdown without logging in.
	
	The UAC has more settings than what the slider shows.
	
	Many of the settings this script reports can be controlled by using Group Policy.
	Windows Settings\Securty Settings\Local Policies\Security Options.
	
	.PARAMETER ComputerName
	This can be:  A Single Computer, The Local Computer, A comma-separated list of computers
	or a Text file list of computers.  For Text files make sure each computer is on a different
	line.
	
	.EXAMPLE
	Get-UACRegistryKey AB-1234
	
	This will provide the current Registry Values for the UAC for the machine AD-1234.
	
	.EXAMPLE
	Get-UACRegistryKey AB-1234, CD-5678, EF-9012
	
	This will provide the current Registry Values for the UAC for each of the machines
	in the comma-separated list.
	
	.EXAMPLE
	Get-UACRegistryKey "C:\My\List\Of\Computers.txt"
	
	This will provide the UAC settings for all computers in the text file list.
	Every computer in the list MUST be on its own line in the text file.
	
	.NOTES
	This will enable and disable the RemoteRegistry service to allow access to remote
	computers.
	
	This must be run as a user that has Administrator rights to the target computers.
#>
param
(
	[Parameter(Mandatory=$False,Position=0)] [string[]]$ComputerName = (Get-ADComputer -Filter {OperatingSystem -like "Windows*"} )
)
$ErrorActionPreference = 'silentlycontinue'	# Skip all the annoying Error Messages

# If the input is a text file, then create a list of entries that are in the text file.
# Else the list will contain a single computer that was input by the user or a list
# of computers that are input separate by commas.
If($ComputerName -like "*.txt")
{
	$CompList = Get-Content ([regex]::matches($ComputerName,'[^\"]+') | %{$_.value})
}
Else
{
	$CompList = $ComputerName
}

#########################################################################################
# VARIABLES
# $Export is a boolean value to tell the script to export information to a csv file or not.
# $UACRegPath is the location in the Registry of the UAC Settings.
# $UACOutput is the output location for the csv file.
#########################################################################################
[boolean]$Export = $False
$UACRegPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System"
$UACOutput = <PATH TO OUTPUT>.csv
#########################################################################################

# Go through each computer in the list and run the operations on it.
Foreach($Computer in $CompList)
{
	# If the computer is offline then skip it.
	If(!(Test-Connection -ComputerName $Computer -TimeToLive 18 -Count 1))
	{
		Write-Host "$Computer - Offline" -Foreground red -Background black
		Write-Host ""
		
		# Output a line for the offline computer to the csv file.
		If($Export -eq $True)
		{
			# Output for an offline PC.
			$objUAC = New-Object PSObject
			$objUAC | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer
			$objUAC | Add-Member -MemberType NoteProperty -Name FilterAdministratorToken -Value "Offline"
			$objUAC | Add-Member -MemberType NoteProperty -Name EnableUIADesktopToggle -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name ConsentPromptBehaviorAdmin -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name ConsentPromptBehaviorUser -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name EnableInstallerDetection -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name ValidateAdminCodeSignatures -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name EnableSecureUIAPaths -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name EnableLUA -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name PromptOnSecureDesktop -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name EnableVirtualization -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name LegalNoticeCaption -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name LegalNoticeText -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name DontDisplayLastUserName -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name SCForceOption -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name ShutdownWithoutLogon -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name UndockWithoutLogon -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name DisableCAD -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name DSCAutomationHostEnabled -Value ""
			$objUAC | Add-Member -MemberType NoteProperty -Name LogonType -Value ""
			
			Export-CSV -Path $UACOutput -InputObject $objUAC -Append -Force
		}
		continue
	}
	Else
	{
		# Enable the Remote Registry service on the target computer.  This allows keys to read/modified.
		Set-Service -ComputerName $Computer -StartUpType Manual -Status Running `
			-Name RemoteRegistry -DisplayName "Remote Registry"

		# Open the HKEY_LOCAL_MACHINE Hive.
		$HKLMKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBasekey('LocalMachine', "$Computer")
		
		# Access the UAC Registry setting in the Registry.
		# $False means this key is accessed in Read-Only Mode.
		$UACKey = $HKLMKey.OpenSubKey("$UACRegPath", $False)
		
		# User Account Control Registry key setting variables.
		$FilterAdministratorToken = "" # UAC: Admin Approval Mode for the built-in Administrator account
		$ConsentPromptBehaviorAdmin = "" # UAC: Behavior of the elevation prompt for administrators in Admin Approval Mode
		$ConsentPromptBehaviorUser = "" # UAC: Behavior of the elevation prompt for standard users
		$EnableInstallerDetection = "" # UAC: Detect application installations and prompt for elevation
		$EnableLUA = "" # UAC: Run all administrators in Admin Approval Mode
		$EnableSecureUIAPaths = "" # UAC: Only elevate UIAccess applications that are installed in secure locations
		$EnableUIADesktopToggle = "" # UAC: Allow UIAccess applications to prompt for elevation without using the secure desktop
		$EnableVirtualization = "" # UAC: Virtualize file and registry write failures to per-user locations
		$PromptOnSecureDesktop = "" # UAC: Switch to the secure desktop when prompting for elevation
		$ValidateAdminCodeSignatures = "" # UAC: Only elevate executables that are signed and validated
		
		# Other Variables related to computer access and settings.
		$DontDisplayLastUserName = "" # Interactive Logon: Do not display last user name
		$SCForceOption = "" # Interactive Logon: Require Smart Card
		$ShutdownWithoutLogon = "" # Shutdown: Allow system to be shut down without having to log on
		$UndockWithoutLogon = "" # Devices: Allow undock without having to log on
		$DisableCAD = "" # Interactive Logon: Do not require CTRL+ALT+DEL
		$DSCAutomationHostEnabled = "" # Desired State Configuration Hose Enabled. Applies for PC's with WMF 5.0.
		$LogonType = "" # Windows Welcome Screen.

		Switch($UACKey.GetValue("ConsentPromptBehaviorAdmin"))
		{
			0{$ConsentPromptBehaviorAdmin = "0 - Elevate Without Prompting"}
			1{$ConsentPromptBehaviorAdmin = "1 - Prompt for Credentials on the Secure Desktop"}
			2{$ConsentPromptBehaviorAdmin = "2 - Prompt for Consent on the Secure Desktop"}
			3{$ConsentPromptBehaviorAdmin = "3 - Prompt for Credentials"}
			4{$ConsentPromptBehaviorAdmin = "4 - Prompt for Consent"}
			5{$ConsentPromptBehaviorAdmin = "5 - Prompt for Consent for Non-Windows Binaries"}
		}
		
		Switch($UACKey.GetValue("ConsentPromptBehaviorUser"))
		{
			0{$ConsentPromptBehaviorUser = "0 - Automatically Deny Elevation Requests"}
			1{$ConsentPromptBehaviorUser = "1 - Prompt for Credentials on the Secure Desktop"}
			3{$ConsentPromptBehaviorUser = "3 - Prompt for Credentials"}
		}
		
		Switch($UACKey.GetValue("EnableInstallerDetection"))
		{
			0{$EnableInstallerDetection = "0 - Disabled"}
			1{$EnableInstallerDetection = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("EnableLUA"))
		{
			0{$EnableLUA = "0 - Disabled"}
			1{$EnableLUA = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("EnableSecureUIAPaths"))
		{
			0{$EnableSecureUIAPaths = "0 - Disabled"}
			1{$EnableSecureUIAPaths = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("EnableUIADesktopToggle"))
		{
			0{$EnableUIADesktopToggle = "0 - Disabled"}
			1{$EnableUIADesktopToggle = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("EnableVirtualization"))	
		{
			0{$EnableVirtualization = "0 - Disabled"}
			1{$EnableVirtualization = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("PromptOnSecureDesktop"))
		{
			0{$PromptOnSecureDesktop = "0 - Disabled"}
			1{$PromptOnSecureDesktop = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("ValidateAdminCodeSignatures"))
		{
			0{$ValidateAdminCodeSignatures = "0 - Disabled"}
			1{$ValidateAdminCodeSignatures = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("dontdisplaylastusername"))
		{
			0{$DontDisplayLastUserName = "0 - Disabled"}
			1{$DontDisplayLastUserName = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("scforceoption"))
		{
			0{$SCForceOption = "0 - Disabled"}
			1{$SCForceOption = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("shutdownwithoutlogon"))
		{
			0{$ShutdownWithoutLogon = "0 - Disabled"}
			1{$ShutdownWithoutLogon = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("undockwithoutlogon"))
		{
			0{$UndockWithoutLogon = "0 - Disabled"}
			1{$UndockWithoutLogon = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("FilterAdministratorToken"))
		{
			0{$FilterAdministratorToken = "0 - Disabled"}
			1{$FilterAdministratorToken = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("DisableCAD"))
		{
			0{$DisableCAD = "0 - Disabled"}
			1{$DisableCAD = "1 - Enabled"}
		}
		
		Switch($UACKey.GetValue("DSCAutomationHostEnabled"))
		{
			0{$DSCAutomationHostEnabled = "0 - Disable Configuring the Machine at Boot-up"}
			1{$DSCAutomationHostEnabled = "1 - Enable Configuring the Machine at Boot-up"}
			2{$DSCAutomationHostEnabled = "2 - Enable Configuring the Machine only if DSC is in pending or current state."}
		}
		
		Switch($UACKey.GetValue("LogonType"))
		{
			0{$LogonType = "0 - Disabled"}
			1{$LogonType = "1 - Enabled"}
		}
		
		# Create an object to store the information.  This will be output to the screen and to
		# a csv file if $Export is set to $True.
		$objUAC = New-Object PSObject
		$objUAC | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer
		$objUAC | Add-Member -MemberType NoteProperty -Name FilterAdministratorToken -Value $FilterAdministratorToken
		$objUAC | Add-Member -MemberType NoteProperty -Name EnableUIADesktopToggle -Value $EnableUIADesktopToggle
		$objUAC | Add-Member -MemberType NoteProperty -Name ConsentPromptBehaviorAdmin -Value $ConsentPromptBehaviorAdmin
		$objUAC | Add-Member -MemberType NoteProperty -Name ConsentPromptBehaviorUser -Value $ConsentPromptBehaviorUser
		$objUAC | Add-Member -MemberType NoteProperty -Name EnableInstallerDetection -Value $EnableInstallerDetection
		$objUAC | Add-Member -MemberType NoteProperty -Name ValidateAdminCodeSignatures -Value $ValidateAdminCodeSignatures
		$objUAC | Add-Member -MemberType NoteProperty -Name EnableSecureUIAPaths -Value $EnableSecureUIAPaths
		$objUAC | Add-Member -MemberType NoteProperty -Name EnableLUA -Value $EnableLUA
		$objUAC | Add-Member -MemberType NoteProperty -Name PromptOnSecureDesktop -Value $PromptOnSecureDesktop
		$objUAC | Add-Member -MemberType NoteProperty -Name EnableVirtualization -Value $EnableVirtualization
		$objUAC | Add-Member -MemberType NoteProperty -Name LegalNoticeCaption -Value $UACKey.GetValue("legalnoticecaption")
		$objUAC | Add-Member -MemberType NoteProperty -Name LegalNoticeText -Value $UACKey.GetValue("legalnoticetext")
		$objUAC | Add-Member -MemberType NoteProperty -Name DontDisplayLastUserName -Value $DontDisplayLastUserName
		$objUAC | Add-Member -MemberType NoteProperty -Name SCForceOption -Value $SCForceOption
		$objUAC | Add-Member -MemberType NoteProperty -Name ShutdownWithoutLogon -Value $ShutdownWithoutLogon
		$objUAC | Add-Member -MemberType NoteProperty -Name UndockWithoutLogon -Value $UndockWithoutLogon
		$objUAC | Add-Member -MemberType NoteProperty -Name DisableCAD -Value $DisableCAD
		$objUAC | Add-Member -MemberType NoteProperty -Name DSCAutomationHostEnabled -Value $DSCAutomationHostEnabled
		$objUAC | Add-Member -MemberType NoteProperty -Name LogonType -Value $LogonType
		
		# Disable the Remote Registry service.
		(Get-Service -ComputerName $Computer -Name RemoteRegistry).stop()
		Set-Service -ComputerName $Computer -StartUpType Disabled -Name RemoteRegistry
		Write-Host ""
		
		$objUAC # Output to Screen.
		
		# Output to csv file if desired. 
		If($Export -eq $True)
		{
			Export-CSV -Path $UACOutput -InputObject $objUAC -Append -Force
		}
	}
}
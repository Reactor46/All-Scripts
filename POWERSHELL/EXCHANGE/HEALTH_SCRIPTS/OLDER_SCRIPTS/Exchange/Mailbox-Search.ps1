

#------------------------------------------------------------------------------
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
#------------------------------------------------------------------------------
# Author: Eyal Doron 
# Version: 1.1 
# Last Modified Date: 30/08/2017  
# Last Modified By: eyal@o365info.com
# Web site: HTTP://o365info.com
# Article name: Search + Save a copy of mail items using PowerShell | Part 2#5
# Article URL: http://o365info.com/search-save-a-copy-of-mail-items-using-powershell-2-5/
# sn -AA001676789
#------------------------------------------------------------------------------
# Hope that you enjoy it ! 
# And, may the force of PowerShell will be with you  :-)
# ------------------------------------------------------------------------------


#------------------------------------------------------------------------------
# PowerShell Functions
#------------------------------------------------------------------------------

function checkConnection ()
{
		$a = Get-PSSession
		if ($a.ConfigurationName -ne "Microsoft.Exchange")
		{
			
			write-host     'You are not connected to Exchange Online PowerShell ;-(         ' 
			write-host      'Please connect using the Menu option 1) Login to Office 365 + Exchange Online using Remote PowerShell        '
			#Read-Host "Press Enter to continue..."
			Add-Type -AssemblyName System.Windows.Forms
			[System.Windows.Forms.MessageBox]::Show("You are not connected to Exchange Online PowerShell ;-( `nSelect menu 1 to connect `nPress OK to continue...", "o365info.com PowerShell script", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			Clear-Host
			break
		}
}



Function Disconnect-ExchangeOnline 
{
Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"} | Remove-PSSession

}


Function Set-AlternatingRows {
       <#
       
       #>
    [CmdletBinding()]
       Param(
             [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [string]$Line,
       
           [Parameter(Mandatory=$True)]
             [string]$CSSEvenClass,
       
        [Parameter(Mandatory=$True)]
           [string]$CSSOddClass
       )
       Begin {
             $ClassName = $CSSEvenClass
       }
       Process {
             If ($Line.Contains("<tr>"))
             {      $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
                    If ($ClassName -eq $CSSEvenClass)
                    {      $ClassName = $CSSOddClass
                    }
                    Else
                    {      $ClassName = $CSSEvenClass
                    }
             }
             Return $Line
       }
}




#------------------------------------------------------------------------------
# Genral
#------------------------------------------------------------------------------
$FormatEnumerationLimit = -1
$Date = Get-Date
$Datef = Get-Date -Format "\Da\te dd-MM-yyyy \Ti\me H-mm" 
#------------------------------------------------------------------------------
# PowerShell console window Style
#------------------------------------------------------------------------------

$pshost = get-host
$pswindow = $pshost.ui.rawui

	$newsize = $pswindow.buffersize
	
	if($newsize.height){
		$newsize.height = 3000
		$newsize.width = 150
		$pswindow.buffersize = $newsize
	}

	$newsize = $pswindow.windowsize
	if($newsize.height){
		$newsize.height = 50
		$newsize.width = 150
		$pswindow.windowsize = $newsize
	}

#------------------------------------------------------------------------------
# HTML Style start 
#------------------------------------------------------------------------------
$Header = @"
<style>
Body{font-family:segoe ui,arial;color:black; }

H1 {font-size: 26px; font-weight:bold;width: 70% text-transform: uppercase; color: #0000A0; background:#2F5496 ; color: #ffffff; padding: 10px 10px 10px 10px ; border: 3px solid #00B0F0;}
H2{ background:#F2F2F2 ; padding: 10px 10px 10px 10px ; color: #013366; margin-top:35px;margin-bottom:25px;font-size: 22px;padding:5px 15px 5px 10px; }

.TextStyle {font-size: 26px; font-weight:bold ; color:black; }

TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 5px;border-style: solid;border-color: #d1d3d4;background-color:#0072c6 ;color:white;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}

.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }



.o365info {height: 90px;padding-top:5px;padding-bottom:5px;margin-top:20px;margin-bottom:20px;border-top: 3px dashed #002060;border-bottom: 3px dashed #002060;background: #00CCFF;font-size: 120%;font-weight:bold;background:#00CCFF url(http://o365info.com/wp-content/files/PowerShell-Images/o365info120.png) no-repeat 680px -5px;
}

</style>

"@

$EndReport = "<div class=o365info>  This report was created by using <a href= http://o365info.com target=_blank>o365info.com</a> PowerShell script </div>"
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------




#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#  Exchange Online Search and copy mail items using Search-Mailbox | PowerShell - Script menu
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

$Loop = $true
While ($Loop)
{
    write-host 
    write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host   "Exchange Online Search and copy mail items using Search-Mailbox | PowerShell - Script menu"
    write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    write-host
	write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor DarkCyan     'Connect Exchange Online using Remote PowerShell        ' 
    write-host -ForegroundColor white  '----------------------------------------------------------------------------------------------' 
	write-host -ForegroundColor Yellow ' 1) Login to Exchange Online using Remote PowerShell ' 
    write-host
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION 0: Assign the required permissions for the PowerShell cmdlet – Search-Mailbox         ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------'
	write-host                                              ' 2) Assign the required permissions to Office 365 Global Administrator  '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION A: Search and copy ALL mail items| Filter - NO Filter         ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------'
	write-host                                              ' 3) Search + Save a copy ALL mail items | Filter scope - NO Filter (SearchQuery)  '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION B: Search and copy mail items| Filter - By mail item type        ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host                                              ' 4)  Search + Save a copy of mail items | Filter scope - Calendar items '
	write-host                                              ' 5)  Search + Save a copy of mail items | Filter scope - Contacts items  '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION C: Search and copy mail items| Filter - By Text String        ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
	write-host                                              ' 6)  Search + Save a copy of mail items | Filter scope - Emails with Text String in mail SUBJECT  '
	write-host                                              ' 7)  Search + Save a copy of mail items | Filter scope - Emails with Text String in mail BODY  '
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION D: Search and copy mail items| Filter - By Date        ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
	write-host                                              ' 8)  Search + Save a copy of mail items | Filter scope - Emails SENT in a Specific date  '
	write-host                                              ' 9)  Search + Save a copy of mail items | Filter scope - Emails SENT in a specific Date Range  '
	write-host                                              ' 10) Search + Save a copy of mail items | Filter scope - Emails RECEIVED in a spesfic Date Range  '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION E: Search and copy mail items| Filter - By Sender and Recipient         ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------'
	write-host                                              ' 11) Search + Save a copy of mail items | Filter scope - Emails sent from specific SENDER  '
	write-host                                              ' 12) Search + Save a copy of mail items | Filter scope - Emails sent TO specific RECIPIENT   '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION F: Search and copy mail items| Filter - By Attachment         ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------'
	write-host                                              ' 13) Search + Save a copy of mail items | Filter scope - Emails that include a specific Attachment file  '
	write-host                                              ' 14) Search + Save a copy of mail items | Filter scope - Emails with Specific Attachment type (suffix)  '
	write-host                                              ' 15) Search + Save a copy of mail items | Filter scope - Emails with Attachment   '
	write-host -ForegroundColor green '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Blue   'SECTION G: Search and copy mail items| Filter - Additional          ' 
    write-host -ForegroundColor green '----------------------------------------------------------------------------------------------'
	write-host                                              ' 16) Search + Save a copy of mail items | Filter scope - E-mail items size greater than X MB  '
	write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
    write-host -ForegroundColor white  -BackgroundColor Magenta 'End of PowerShell - Script menu ' 
    write-host -ForegroundColor green  '----------------------------------------------------------------------------------------------' 
	
	
	write-host -ForegroundColor Red            "20)  Disconnect PowerShell session" 
    write-host
    write-host -ForegroundColor Red            "21)  Exit (Or use the keyboard combination - CTRL + C)" 
    write-host
    $opt = Read-Host "Select an option [1-21]"
    write-host $opt
    switch ($opt) 


{


	
1
{

#####################################################################
# Connect Exchange Online using Remote PowerShell
#####################################################################

# == Section: General information ===

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------                                                                
write-host  -ForegroundColor white  	'To be able to use the PowerShell menus in the script,  '
write-host  -ForegroundColor white  	'you will need to Login to Exchange Online using Remote PowerShell. '
write-host  -ForegroundColor white  	'In the credentials windows that appear,   '
write-host  -ForegroundColor white  	'provide your Office 365 Global Administrator credentials.  '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'$UserCredential = Get-Credential '
write-host  -ForegroundColor Cyan    	'$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection '
write-host  -ForegroundColor Cyan    	'Import-PSSession $Session '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

Disconnect-ExchangeOnline 
# Specify your administrative user credentials on the line below 

$user = “Provide credentials”

$UserCredential = Get-Credential -Credential $user

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session

}


#=========================================================================================
# SECTION 0: Assign the required permissions for the PowerShell cmdlet – Search-Mailbox
#=========================================================================================

2
{


#####################################################################
# Provide the required permissions to Global Administrator 
#####################################################################

checkConnection
# General information 

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		----------------------------------------------------------------------------   
write-host  -ForegroundColor white  	' Provide the required permissions to Global Administrator  '
write-host  -ForegroundColor white		----------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'This option will: '
write-host  -ForegroundColor white  	'By default, Office 365 Global Administrator DONT HAVE the required permissions to '
write-host  -ForegroundColor white  	'activate the PowerShell cmdlet Search-Mailbox which is used '
write-host  -ForegroundColor white  	'to perform a search in Exchange Online mailbox or for deleting mail items.'
write-host  -ForegroundColor white  	'To be able to run the required PowerShell cmdlet Search-Mailbox, we need to assign 365 Global Administrator '
write-host  -ForegroundColor white  	'the required permissions by adding his name as a member in the Discovery Management security group. '
write-host  -ForegroundColor white  	'Add the Mailbox Import Export role to Discovery Management group Discovery Management group. '
write-host  -ForegroundColor white		---------------------------------------------------------------------------- 
write-host  -ForegroundColor white  	'Note - in case that you try to run the PowerShell script menu without the required permissions, '
write-host  -ForegroundColor white  	'you will get an error such as: '
write-host  -ForegroundColor red  	    'Search-Mailbox - The term Search-Mailbox is not recognized as the name of a cmdlet,script file, or operable program '
write-host  -ForegroundColor Yellow  	'or '
write-host  -ForegroundColor red  	    'A parameter cannot be found that matches parameter name DeleteContent. '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Add-RoleGroupMember -Identity "Discovery Management" -Member <Identity> '
write-host  -ForegroundColor Cyan    	'New-ManagementRoleAssignment  -SecurityGroup "Discovery Management" -Role "Mailbox Import Export" '
write-host  -ForegroundColor white		--------------------------------------------------------------------  
write-host  -ForegroundColor Cyan    	'Note – in an Office 365 environment, the permissions update process  '
write-host  -ForegroundColor Cyan    	'could take half an hour or several hours. '
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# User input

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 1 parameter:"  
write-host
write-host -ForegroundColor Yellow	"1.  Office 365 Global Administrator  "  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Office 365 Global Administrator name" 
write-host -ForegroundColor Yellow	"NOTE - The Global Administrator name that you provide should be the user name "   
write-host -ForegroundColor Yellow	"which you use for creating the current remote PowerShell session. "     
write-host -ForegroundColor Yellow	"For example:  Bobadmin@o365info.com"
write-host
$Alias  = Read-Host "Type the Office 365 Global Administrator name"
write-host
write-host


# Section 1#2 – Add the Global Administrator account to the Discovery Management group
# Display Information 

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue   "Display information about: Discovery Management Group members => BEFORE the update"
write-host -------------------------------------------------------------------------------------------------
Get-RoleGroupMember -Identity "Discovery Management"  | Out-String
write-host -------------------------------------------------------------------------------------------------


write-host
write-host  -ForegroundColor Cyan  "Step 1#2 - Add the Global Administrator account to the Discovery Management group"

# PowerShell Command
Add-RoleGroupMember -Identity "Discovery Management" -Member $Alias

# Display Information 
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  "Display information about: Discovery Management Group members  => AFTER the update"
write-host -------------------------------------------------------------------------------------------------
Get-RoleGroupMember -Identity "Discovery Management"  | Out-String
write-host -------------------------------------------------------------------------------------------------

# Section 2#2 – Add the role Mailbox Import Export to the Discovery Management group

write-host
write-host  -ForegroundColor Cyan  "Step 2#2 - Add the role Mailbox Import Export to the Discovery Management group"

# Add the Mailbox Import Export role to Discovery Management group
New-ManagementRoleAssignment  -SecurityGroup "Discovery Management" -Role "Mailbox Import Export"

#Display Information 
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Magenta  "Display information about: the Discovery Management Role and permissions assignment  => AFTER the update"
write-host -------------------------------------------------------------------------------------------------
Get-RoleGroup "Discovery Management" | Select-Object  -ExpandProperty RoleAssignments     | Out-String
write-host -------------------------------------------------------------------------------------------------

# End the menu command
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}





#===============================================================================================================
# SECTION A: Search and copy ALL mail items| Filter - NO Filter
#===============================================================================================================

		
3
{


##################################################################################################################
# Search + Save a copy ALL mail items | Filter scope - NO Filter (SearchQuery)
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy ALL mail items | Filter scope - NO Filter (SearchQuery) '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. DONT Filter the search result - Search ALL mail items (BACKUP the Mailbox)  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 2 parameters:"  
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host

$TargetFolderName = "Search Results $SourceMailBox" 


# == Display information about search query parameters

write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox  -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	




#===============================================================================================================
# SECTION B: Search and copy mail items| Filter - By mail item type 
#===============================================================================================================


4
{


########################################################################################
# Search + Save a copy of mail items | Filter scope - Calendar items
########################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Calendar items '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only Calendar items  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  "Kind:meetings" -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 2 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. Mail Items type - Calendar 
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchDumpsterOnly -SearchQuery  "Kind:meetings" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}




5
{


########################################################################################
# Search + Save a copy of mail items | Filter scope - Contacts items
########################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Contacts items '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only Cotacts items  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  "Kind:contacts" -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 2 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host


$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. Mail Items type - Contacts 
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  "Kind:contacts" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}

			
#===============================================================================================================
# SECTION C: Search and copy mail items| Filter - By Text String
#===============================================================================================================
						

6
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails with Text String in mail SUBJECT
##################################################################################################################

checkConnection
# General information  

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails with Text String in mail SUBJECT '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items with spesfic Text String in mail subject '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  'Subject:"<Text String>"' -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host

write-host
write-host -ForegroundColor Yellow	"3.  Text String " 
write-host -ForegroundColor Yellow	"The term - Text String, define the TEXT taht we look for in the mail SUBJECT" 
write-host -ForegroundColor Yellow	"Provide the Text String"    
write-host -ForegroundColor Yellow	"For example:  Metting with John"
write-host
$TextString = Read-Host "Type the Text String we look for in the mail SUBJECT"
write-host
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Text String we look for in the mail SUBJECT is  - $TextString
write-host -------------------------------------------------------------------------------------------------


# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  Subject:"$TextString" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}

							

7
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails with Text String in mail BODY 
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails with Text String in mail BODY  '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items with spesfic Text String in mail BODY '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  'Body:"<Text String>"' -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host

write-host -ForegroundColor Yellow	"3.  Text String " 
write-host -ForegroundColor Yellow	"The term - Text String, define the TEXT taht we look for in the mail BODY" 
write-host -ForegroundColor Yellow	"Provide the Text String"    
write-host -ForegroundColor Yellow	"For example:  Metting with John"
write-host
$TextString = Read-Host "Type the Text String we look for in the mail BODY"
write-host
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Text String we look for in the mail BODY is  - $TextString
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery   body:"$TextString" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}



	
#===============================================================================================================
# SECTION D: Search and copy mail items| Filter - By date
#===============================================================================================================


8
{


##################################################################################################################
#  Search + Save a copy of mail items | Filter scope - Emails from a Specific date
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails from a Specific date '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items  Mail items sent in a specific date  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  sent:mm/dd/yyyy  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host
write-host -ForegroundColor Yellow	"3.  The specific date " 
write-host -ForegroundColor Yellow	"spesficy the date useing the follwoing date format mm/dd/yyyy" 
write-host -ForegroundColor Yellow	"For example:  07/25/2017"
write-host
$DateString = Read-Host "Type the specific date"
write-host
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Mail flow is                          - Sent email 
write-host 5. The Specific date the we look for is      - $DateString
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  sent:$DateString -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}

			
9
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails SENT in a specific Date Range
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails SENT in a specific Date Range '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items  Mail items sent in a spesfic Date Range  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  {sent:mm/dd/yyyy..mm/dd/yyyy}  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 4 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host
write-host -ForegroundColor Yellow	"3.  The START date " 
write-host -ForegroundColor Yellow	"Spesficy the start date use the follwoing format mm/dd/yyyy" 
write-host -ForegroundColor Yellow	"For example:  10/07/2017"
write-host
$DateStart = Read-Host "Type the specific START date"
write-host
write-host
write-host
write-host -ForegroundColor Yellow	"4.  The END date " 
write-host -ForegroundColor Yellow	"Spesficy the end date use the follwoing format mm/dd/yyyy" 
write-host -ForegroundColor Yellow	"For example:  25/07/2017"
write-host
$DateEnd = Read-Host "Type the specific END date"
write-host
write-host
$TargetFolderName = "Search Results $SourceMailBox" 

write-host -ForegroundColor Yellow	"In case that you get an error such as - The KQL parser threw an exception.  " 
write-host -ForegroundColor Yellow	"Use the mount name instead of the number  " 
write-host -ForegroundColor Yellow	"For example- 02/July/2017 " 
write-host -ForegroundColor Yellow	"https://support.microsoft.com/en-us/help/2982852/ediscovery-search-error-when-you-use-kql-format-for-dates-in-exchange " 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Mail flow is          - Sent email 
write-host 5. The START date is         - $DateStart 
write-host 6. The END date is           - $DateEnd 
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  {sent:$DateStart..$DateEnd} -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	



			
10
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails RECEIVED in a spesfic Date Range
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails RECEIVED in a spesfic Date Range'
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items  Mail items sent in a specific date  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  {Received:mm/dd/yyyy..mm/dd/yyyy}  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 4 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host
write-host -ForegroundColor Yellow	"3.  The START date " 
write-host -ForegroundColor Yellow	"Spesficy the start date use the follwoing format mm/dd/yyyy" 
write-host -ForegroundColor Yellow	"For example:  10/07/2017"
write-host
$DateStart = Read-Host "Type the specific START date"
write-host
write-host
write-host
write-host -ForegroundColor Yellow	"4.  The END date " 
write-host -ForegroundColor Yellow	"Spesficy the end date use the follwoing format mm/dd/yyyy" 
write-host -ForegroundColor Yellow	"For example:  25/07/2017"
write-host
$DateEnd = Read-Host "Type the specific END date"
write-host
write-host
$TargetFolderName = "Search Results $SourceMailBox" 

write-host -ForegroundColor Yellow	"In case that you get an error such as - The KQL parser threw an exception.  " 
write-host -ForegroundColor Yellow	"Use the mount name instead of the number  " 
write-host -ForegroundColor Yellow	"For example- 02/July/2017 " 
write-host -ForegroundColor Yellow	"https://support.microsoft.com/en-us/help/2982852/ediscovery-search-error-when-you-use-kql-format-for-dates-in-exchange " 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Mail flow is          - Received email 
write-host 5. The START date is         - $DateStart 
write-host 6. The END date is           - $DateEnd 
write-host -------------------------------------------------------------------------------------------------


# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  {Received:$DateStart..$DateEnd} -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	



	
#===============================================================================================================
# SECTION E: Search and copy mail items| Filter - By sender and Recipient
#===============================================================================================================

		
11
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails sent from specific SENDER 
##################################################################################################################
checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	' Search + Save a copy of mail items | Filter scope - Emails sent from specific SENDER  '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items -Emails  sent by a specific  SENDER  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery from:"<E-mail address>"  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host -ForegroundColor Yellow	"3.  SENDER idintity " 
write-host -ForegroundColor Yellow	"Spesficy the SENDER idintity using the Email Adtress" 
write-host -ForegroundColor Yellow	"For example:  Alice@contoso.com"
write-host
$SenderAddress = Read-Host "Type the SENDER idintity"
write-host
write-host
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. Filtered search result by - Sender 
write-host 5. The Sender idintity is    - $SenderAddress 
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  From:"$SenderAddress" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	


		
12
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails sent TO specific RECIPIENT
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails sent TO specific RECIPIENT '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items -Emails  sent TO a specific RECIPIENT  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  to:"<E-mail address>"  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host
write-host -ForegroundColor Yellow	"3.  RECIPIENT idintity " 
write-host -ForegroundColor Yellow	"Spesficy the RECIPIENT idintity by using the Email Adtress" 
write-host -ForegroundColor Yellow	"For example:  Alice@contoso.com"
write-host
$RecipientAddress = Read-Host "Type the RECIPIENT idintity"
write-host
write-host
write-host


$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. Filtered search result by - Sender 
write-host 5. The Recipient idintity is - $RecipientAddress
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  to:"$SenderAddress" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	

#===============================================================================================================
# SECTION F: Search and copy mail items| Filter - By Attachment 
#===============================================================================================================
		
13
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails that include a specific Attachment file
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails that include a specific Attachment file '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items that include a specific attachment file  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  attachment:"<Attachment file name>"  -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host


# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host
write-host -ForegroundColor Yellow	"3.  Attachment File name " 
write-host -ForegroundColor Yellow	"Spesficy the Attachment File name" 
write-host -ForegroundColor Yellow	"For example:  report.pdf"
write-host
$AttachmentFile = Read-Host "Type the Attachment File name"
write-host
write-host
write-host


$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Attachment File name is - $AttachmentFile
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  attachment:"$AttachmentFile" -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	

		
14
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails with Specific attachment type (suffix)
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails with Specific Attachment type (suffix) '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items that have a specific attachment type (suffix)  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  {Attachment -like "*.<suffix>"}   -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host -ForegroundColor Yellow	"3.  Attachment File type (suffix) " 
write-host -ForegroundColor Yellow	"Spesficy the Attachment File type (suffix)" 
write-host -ForegroundColor Yellow	"For example:  csv"
write-host
$Attachmentsuffix = Read-Host "Type the Attachment File type (suffix)"
write-host
write-host
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Attachment File type(suffix) is - $Attachmentsuffix
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  {Attachment -like "*.$Attachmentsuffix"} -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	
		
15
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - Emails with attachment
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - Emails with attachment '
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items with attachment  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery {HasAttachment -eq $true} -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 2 parameters:"  
write-host
wwrite-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host

$TargetFolderName = "Search Results $SourceMailBox" 

# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. Email items with attachment          - Yes
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  {HasAttachment -eq $true} -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full

# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}


#===============================================================================================================
# SECTION G: Search and copy mail items| Filter -  Additional 
#===============================================================================================================

		
16
{


##################################################################################################################
# Search + Save a copy of mail items | Filter scope - E-mail items size greater than X MB
##################################################################################################################

checkConnection
# General information 

clear-host

write-host
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                         
write-host  -ForegroundColor white		Information                                                                                          
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------                                                           
write-host  -ForegroundColor white  	'Search + Save a copy of mail items | Filter scope - E-mail items size greater than X MB'
write-host  -ForegroundColor white  	'This menu option will perform the following tasks: '
write-host  -ForegroundColor white  	'1. Search for mail items in a specific Mailbox (including Recovery folder)'
write-host  -ForegroundColor white  	'2. Copy the mail items to a target Mailbox (saved in target folder)'
write-host  -ForegroundColor white  	'3. Create a Target Folder - NEW folder in the Target Mailbox taht will be named by using the name of the Source Mailbox '
write-host  -ForegroundColor Magenta  	'4. Filter the search result - Search only mail items that E-mail items size greater then X MB  '
write-host  -ForegroundColor white  	'5. Create a LOG file that include information about - each Email item, that was copied to the Target mailbox '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'NOTE - The Search-Mailbox cmdlet returns up to 10000 results per mailbox if a search query is specified. '
write-host  -ForegroundColor white		------------------------------------------------------------------------------------------------------------------   
write-host  -ForegroundColor white  	'The PowerShell command that we use is: '
write-host  -ForegroundColor Cyan    	'Search-Mailbox <Source Mailbox> -SearchQuery  {Size -gt <size in KB or MB>} -TargetMailbox <Target mailbox> -TargetFolder <Target Folder> -LogLevel Full '
write-host
write-host  -ForegroundColor Magenta	oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo                                          
write-host
write-host

# == Section: User input ===

write-host -ForegroundColor white	'User input '
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
write-host -ForegroundColor Yellow	"You will need to provide 3 parameters:"  
write-host
write-host
write-host -ForegroundColor Yellow	"1. Source mailbox" 
write-host -ForegroundColor Yellow	"The term Source mailbox, define the mailbox where the information (Mail Items) is searched"  
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Source mailbox"     
write-host -ForegroundColor Yellow	"For example:  John@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$SourceMailBox = Read-Host "Type the Source mailbox name"

write-host
write-host
write-host -ForegroundColor Yellow	"2. Target mailbox" 
write-host -ForegroundColor Yellow	"The term Target mailbox, define the mailbox which will store the search results (a copy of the Mail Items)" 
write-host -ForegroundColor Yellow	"Provide the Identity (Alias or E-mail address) of the Target mailbox"    
write-host -ForegroundColor Yellow	"For example:  Admin@o365info.com"
write-host
write-host -ForegroundColor white	---------------------------------------------------------------------------- 
$TargetMailBox = Read-Host "Type the Destination (Target) mailbox name"
write-host
write-host -ForegroundColor Yellow	"3.  Email item size in Megabyte " 
write-host -ForegroundColor Yellow	"Spesficy the Email size in Megabyte" 
write-host -ForegroundColor Yellow	"For example:  1MB"
write-host
$EmailSize = Read-Host "Type the Email size in Megabyte"
write-host
write-host
write-host


$TargetFolderName = "Search Results $SourceMailBox" 
# == Display information about search query parameters
write-host
write-host -------------------------------------------------------------------------------------------------
write-host -ForegroundColor white  -BackgroundColor Blue  information about the serach Task
write-host -------------------------------------------------------------------------------------------------
write-host The search mailbox task will be implemented by using the following parameters:
write-host
write-host 1. The Source mailbox is               – $SourceMailBox 
write-host 2. The Target mailbox is               - $TargetMailBox 
write-host 3. The Target Folder name is           - $TargetFolderName
write-host
write-host -ForegroundColor white  -BackgroundColor DarkCyan Filtered scope
write-host -------------------------------------------------------------------------------------------------
write-host 4. The Email ietms size is grater then - $EmailSize
write-host -------------------------------------------------------------------------------------------------

# == Section: PowerShell Command  ===	

Search-Mailbox $SourceMailBox -SearchQuery  {Size -gt $EmailSize} -TargetMailbox $TargetMailBox -TargetFolder $TargetFolderName -LogLevel Full


# == Section: End the menu command ===	
write-host
write-host
Read-Host "Press Enter to continue..."
write-host
write-host

}
	




	

	
#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
# Section END - Disconnect PowerShell session 
#<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


	
20
{

##########################################
# Disconnect PowerShell session  
##########################################


write-host -ForegroundColor Yellow Choosing this option will Disconnect the current PowerShell session 

Disconnect-ExchangeOnline 

write-host
write-host

#———— Indication ———————

if ($lastexitcode -eq 0)
{
write-host -------------------------------------------------------------
write-host "The command complete successfully !" -ForegroundColor Yellow
write-host "The PowerShell session is disconnected" -ForegroundColor Yellow
write-host -------------------------------------------------------------
}
else

{
write-host "The command Failed :-(" -ForegroundColor red

}

#———— End of Indication ———————


}




21
{

##########################################
# Exit  
##########################################


$Loop = $true
Exit
}

}


}

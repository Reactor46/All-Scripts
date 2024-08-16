<#

.SYNOPSIS
This script reports Retention Tag information for items in chosen mailbox on Exchange 2010

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Version 1.0, 17 January 2013

Serkan Varoglu
Exchange Server MVP

.LINK
Blog:
	http:\\Get-Mailbox.org (Turkish)
	http:\\en.Get-Mailbox.org (English)
Podcast: 
	http:\\www.TheUCArchitects.com
Twitter: 
	@SRKNVRGL

.DESCRIPTION
This script uses EWS to report Retention information for items in chosen mailbox on Exchange 2010. 
Please make sure EWS Settings are correct before running this script.

!! If you have NOT installed EWS Managed API download and install it from: 
http://www.microsoft.com/en-us/download/details.aspx?id=35371

!! This script uses Impersonation. The account you are using for reporting must have impersonation rights on Exchange 2010. 
For more information on Impersonation: 
http://msdn.microsoft.com/en-us/library/exchange/dd633680(v=exchg.80).aspx

!! Exchange 2010 Throttling Policy might cause reporting have missing items. If you are going to use this script against large mailboxes make sure Throttling policy is set correctly.
For more information on Exchange 2010 EWS Throttling please read: 
http://msdn.microsoft.com/en-us/library/exchange/hh881884(v=exchg.140).aspx

.NOTES

.PARAMETER Identity 
Mailbox Alias or SMTP Address
	
.PARAMETER ExpiresInDays
Use this parameter if you want to report items that will expire in X number of days.
!! If you do not use ExportInDays parameter the report will be include all items.

.PARAMETER OnlyTheseFolders
If you input the folder names in this String value, script will report only these folders.
If you want to include subfolders you can use *
Folder Name Format Example: "\Inbox","Calendar","Serkan"

.PARAMETER ExcludeFolders
If you input the folder names in this String value, script will exclude these folders from reporting and calculations.
If you want to include subfolders you can use *
Folder Name Format Example: "\Inbox","Calendar","Serkan"

.PARAMETER HTMLReport
If you want to save the report as an HTML file use this switch. The script will create a report in the following format in the same directory with the script.
SMTPAddress-daymonthhourminute.html
!! Please make sure you have write permissions on the directory that you are running the script.

.PARAMETER CSVReport
If you want to save the report as an CSV file use this switch. The script will create a report in the following format in the same directory with the script.
SMTPAddress-daymonthhourminute.csv
!! Please make sure you have write permissions on the directory that you are running the script.

.PARAMETER HideNeverExpires
This switch will hide all items that has no expiration date from your report.

.PARAMETER EmailUser
Coming Soon
	
.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -HTMLReport

This will create a HTML Report for Serkan user and will include all mailbox items. HTML Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -CSVReport

This will create a CSV Report for Serkan user and will include all mailbox items. CSV Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -CSVReport -HideNeverExpires

This will create a CSV Report for Serkan user and will include all mailbox items. If item does not have any expiration date it will be excluded. CSV Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -OnlyTheseFolders "\Inbox*","Drafts" -CSVReport

This will create a CSV Report for Serkan user and will include Drafts, Inbox and Inbox Subfolders. CSV Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -ExcludeFolders "\Conversation History*","Temp" -CSVReport

This will create a CSV Report for Serkan user and will EXCLUDE Temp, Conversation History and Conversation History Subfolders. CSV Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -ExpiresInDays 60 -CSVReport

This will create a CSV Report for Serkan user for items that will expire in next 60 days. CSV Report will be saved in the same folder as the script.

.EXAMPLE
.\Get-RetentionExpiration -Identity Serkan -ExpiresInDays 30 -CSVReport -HTMLReport

This will create a CSV Report for Serkan user for items that will expire in next 30 days. Both HTML Report and CSV Report will be saved in the same folder as the script.

.NOTES

Special Thanks to Glen Scales - http://gsexdev.blogspot.com/

#>

Param(
[Parameter(Position=1,Mandatory=$True,HelpMessage='Identity (Mailbox Alias)')][String]$Identity,
[Parameter(Position=2,Mandatory=$False,HelpMessage='Expiration Range (Must be between 0-24855)')][ValidateRange(0,24855)][Int]$ExpiresInDays,
[Parameter(Position=3,Mandatory=$False,HelpMessage='HTML Output')][Switch]$HTMLReport,
[Parameter(Position=3,Mandatory=$False,HelpMessage='CSV Output')][Switch]$CSVReport,
[Parameter(Position=4,Mandatory=$False,HelpMessage='Report Only These Folders (Example: "\Inbox*","Calendar","Serkan")')][String[]]$OnlyTheseFolders,
[Parameter(Position=5,Mandatory=$False,HelpMessage='Exclude These Folders from the Report (Example: "\Inbox*","Calendar","Serkan")')][String[]]$ExcludeFolders,
[Parameter(Position=6,Mandatory=$False,HelpMessage='Add if you do not want to report items that will never expire')][Switch]$HideNeverExpires,
[Parameter(Position=7,Mandatory=$False,HelpMessage='Coming Soon...')][Switch]$EMailUser
)
#Set Error Action To SilentlyContinue
$ErrorActionPreference = "SilentlyContinue"
$StartTime = Get-Date
function connect($PrimarySMTPAddress)
{
## Requires the EWS Managed API and Powershell V2.0 or greator  
## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
#$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1 
#$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010 

## Create Exchange Service Object 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
#$creds = New-Object System.Net.NetworkCredential("user@domain.com","password")   
#$service.Credentials = $creds      
  
#Credentials Option 2  
$service.UseDefaultCredentials = $true  
  
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
## end code from http://poshcode.org/624
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($PrimarySMTPAddress,{$true})  
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://FQDN/ews/exchange.asmx"  
#$service.Url = $uri    

## Optional section for Exchange Impersonation  
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$PrimarySMTPAddress);
$service
}

#Define Function to convert String to FolderPath 
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
}  
#Define Function to Check even or odd
Function check-even ($num) {[bool]!($num%2)}

#Check Report Request
if ((!$HTMLReport) -and (!$CSVReport))
{
"You have not chosen to output any report. Please use -HTMLReport or -CSVReport or both in your switch!!!"
Break
}
#Collect Mailbox Retention Information
$MailboxInformation = Get-Mailbox $Identity -ErrorAction SilentlyContinue
$PrimarySMTPAddress = [String]$MailboxInformation.PrimarySMTPAddress
if ($MailboxInformation)
	{
		#Check if Mailbox have Retention Policy Applied.
		if ($MailboxInformation.RetentionPolicy)
			{
			$RetentionPolicyApplied = Get-RetentionPolicy ($MailboxInformation.RetentionPolicy) -ErrorAction SilentlyContinue
			#Retention Policy Tags that exists in applied Retention Policy.
			$RetentionPolicyTagsApplied = Get-RetentionPolicy $RetentionPolicyApplied | select-object -expand RetentionPolicyTagLinks
			#Collect information about Retention Tags on this policy and store in $PolicyTable
			$PolicyTable = @()
			$RetentionPolicyTags = @()
			foreach ($RetentionPolicyTaginPolicy in $RetentionPolicyTagsApplied)
				{
				$RetentionPolicyTags += Get-RetentionPolicyTag $RetentionPolicyTaginPolicy.Name
				}
			foreach ($RetentionPolicyTag in $RetentionPolicyTags)
				{
				$tagValue = New-Object PSobject
				$RetentionPolicyTagByte = (new-Object Guid($RetentionPolicyTag.RetentionId)).ToByteArray()
				$tagValue | add-member -membertype noteproperty -name PolicyName -value $RetentionPolicyTag.Name
				$tagValue | add-member -membertype noteproperty -name PolicyTagByte  -value ([string]$RetentionPolicyTagByte)
				$tagValue | add-member -membertype noteproperty -name MessageClass  -value $RetentionPolicyTag.MessageClass
				$tagValue | add-member -membertype noteproperty -name Action  -value $RetentionPolicyTag.RetentionAction
				$tagValue | add-member -membertype noteproperty -name AgeLimit  -value $RetentionPolicyTag.AgeLimitforRetention
				$tagValue | add-member -membertype noteproperty -name Type  -value $RetentionPolicyTag.Type
				$tagValue | add-member -membertype noteproperty -name MovetoDestinationFolder  -value $RetentionPolicyTag.MovetoDestinationFolder
				$PolicyTable += $tagValue
				}				
			}
		else
			{
			$MailboxInformation.DisplayName + "does not have any Retention Policy!!"
			Break
			}
	}
else
	{
	"Can not find the mailbox. Please try again!!"
	Break
	}

#Default Retention Tag that applies to All folders
$DefaultRetentionTag = $PolicyTable | ?{$_.Type -eq "All"}
	
#HTML Header
if ($HTMLReport)
{
#Save the report as file
$t = Get-Date -UFormat %d%m%H%M
$HTMLReportFile = "$($PrimarySMTPAddress)-$($t).html"
$HTMLHeader = "<html><body><table width=""100%"" cellpadding=""1"" cellspacing=""1""><tr bgcolor=""#66CCFF""><td style=""border:1px solid black;"" width=""20%"" nowrap=""nowrap"">Folder Name</td><td style=""border:1px solid black;"" width=""50%"" nowrap=""nowrap"">Subject</td><td style=""border:1px solid black;"" width=""10%"" nowrap=""nowrap"">Recieved Date</td><td style=""border:1px solid black;"" width=""10%"" nowrap=""nowrap"">Expiration Date</td><td style=""border:1px solid black;"" width=""5%"" nowrap=""nowrap"">Size</td><td style=""border:1px solid black;"" width=""5%"" nowrap=""nowrap"">Retention Tag</td></tr>"
$HTMLHeader | Out-file -Encoding UTF8 $HTMLReportFile
}
#CSV Report
if ($CSVReport)
{
#Save the report as file
$t = Get-Date -UFormat %d%m%H%M
$CSVReportFile = "$($PrimarySMTPAddress)-$($t).csv"
$CSVHeader = "Folder Name, Subject, Recieved Date, Expiration Date, Size, Retention Tag"
$CSVHeader | Out-file -Encoding UTF8 $CSVReportFile
$CSVReportFile
}

cls
#Screen Message
For($s = 1; $s -le 12; $s++){""}
"##########################################################################################"
""
"Processing : " + $MailboxInformation.DisplayName
""
"Retention Policy: " + $RetentionPolicyApplied.Name
""
"##########################################################################################"
""
"Retention Policy Tags available to User"
$PolicyTable | ft PolicyName,Type,AgeLimit,Action
"##########################################################################################"
"Folder Path -- Retention Policy Tag on Folder"
"---------------------------------------------"

#Define Extended properties for Folders
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 

#Folder Type
$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

#Folder Path
$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String); 
$psPropertySet.Add($PR_Folder_Path); 
#PR_POLICY_TAG 0x3019
$PR_POLICY_TAG = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
$psPropertySet.Add($PR_POLICY_TAG);

$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$PrimarySMTPAddress);

#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  

#Deep Transval will ensure all folders in the search path are returned  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
$fvFolderView.PropertySet = $psPropertySet;

#The Search filter will exclude any Search Folders
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  

##Define Extended properties for items
$psItemPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

#PR_RETENTION_DATE 0x301C
$PR_RETENTION_DATE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301C,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
$psItemPropertySet.Add($PR_RETENTION_DATE);

#PR_POLICY_TAG 0x3019
$psItemPropertySet.Add($PR_POLICY_TAG);

#Define the ItemView used for Export should not be any larger then 1000 folders due to throttling
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000);
$ivItemView.PropertySet = $psItemPropertySet;

#Initial Folder Count
$fpathcount = 0

#Create an Empty array for FolderIDs
$fiAllResult = @()
#Connect
$service = connect $PrimarySmtpAddress
#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox 

#If Expiration Date is requested set the filter for it
if ($ExpiresInDays)
{
$ExpirationDate = (get-date).adddays($ExpiresInDays)
$ItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo($PR_RETENTION_DATE,$ExpirationDate); 
}
#Get all folders in the mailbox from Top of Information Store and store information in $fiAllResult
do 
{  
	$fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
	$fvFolderView.Offset += $fiResult.Folders.Count
	$fiAllResult += $fiResult
}
while($fiResult.MoreAvailable -eq $true)
#Start processing each folder
foreach($ffFolder in $fiAllResult)
	{
		$CurrentFolder = $Null
		$fiAllItemCount = 0
		$fiItemResultCount = 0
		$fpathcount++	
		#HTML row color for each folder
		$fpathnum = check-even $fpathcount
		if ($fpathnum -eq $True){$color = """#E0E0E0"""}else{$color = """#F0F0F0"""}
		#Get the folder ID and add it to FolderIDs Array for reporting
		$FolderID = $null
		$FolderID = $ffFolder.ID
		$FolderIDs += $FolderID
		$foldpathval = $null
		#Try to get the Policy Tag Value
		$FolderRetentionTag = $null
		$FolderRetentionTagValue = $null
		$RetentionTagonFolder = $null
		if ($ffFolder.TryGetProperty($PR_POLICY_TAG,[ref] $FolderRetentionTag))  
					{  
						$FolderRetentionTagValue = $FolderRetentionTag
						#Get Folder Retention Policy Tag
						$RetentionTagonFolder = $PolicyTable | ?{$_.PolicyTagByte -eq [String]$FolderRetentionTagValue}
					}
		elseif ($DefaultRetentionTag)
					{
						$RetentionTagonFolder = $DefaultRetentionTag
						$FolderRetentionTagValue = $DefaultRetentionTag.PolicyTagByte
					}
		#Try to get the FolderPath Value and then covert it to a usable String
		if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
					{  
						$binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
						$hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
						$hexString = $hexArr -join ''  
						$hexString = $hexString.Replace("FEFF", "5C00")  
						$fpath = ConvertToString($hexString)  
					}
		
		#Only These Folders
		if ($OnlyTheseFolders)
		{
			foreach ($FilterFolderName in $OnlyTheseFolders)
			{
				if ($fpath -like $FilterFolderName)
				{
				$CurrentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
				}
			}
		}
		else
		{
			$CurrentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderID)
		}
		#Exclude Folders
		if ($ExcludeFolders)
		{
			foreach ($ExcludedFolder in $ExcludeFolders)
			{
				if (($fpath -like $ExcludedFolder))
				{
				$CurrentFolder = $Null
				$fpath + " -- Folder Excluded"
				}
			}
		}
		#Display Which Folder is in process
		Write-Progress -activity "Processing Folder $($fpath)" -status "Status: $($fpathcount) of $($fiAllResult.count) folders" -PercentComplete (($fpathcount / $fiAllResult.count)*100) -id 1
		#Display Folder Retention Tag Information
		
		if ($CurrentFolder)
		{
			if (($DefaultRetentionTag) -and ($RetentionTagonFolder))
			{
				$fpath + " -- " + $RetentionTagonFolder.PolicyName
			}
			elseif ((!$DefaultRetentionTag) -and (!$RetentionTagonFolder))
			{
				$fpath + " -- " + "No Retention Policy Tag"	
			}
			#The Do loop will handle any paging that is required if there are more the 1000 items in the folder 
			do 
			{
				if ($ExpiresInDays)
				{
					$fiItemResult = $service.FindItems($CurrentFolder.Id,$ItemSearchFilter,$ivItemView) 
					$fiItemResultCount += $fiItemResult.Items.count
				}
				else
				{
					$fiItemResult = $service.FindItems($CurrentFolder.Id,$ivItemView)
					$fiItemResultCount += $fiItemResult.Items.count
				}
				$fiItems = $fiItemResult.Items
				#Process each item for Retention Information
				foreach($Item in $fiItems)
				{  
					if ($fiAllItemCount -gt 0)
					{
					Write-progress -Activity "Processing Items" -Status "Status: $($fiAllItemCount) of $($fiItemResultCount)" -percentcomplete (($fiAllItemCount / $fiItemResultCount)*100) -ParentID 1
					}
						$fiAllItemCount++
						$RetentionDate = $null
						$RetentionDateValue = $null
						if ($Item.TryGetProperty($PR_RETENTION_DATE,[ref] $RetentionDate))  
						{  
							$RetentionDateValue = $RetentionDate
						} 
						$RetentionTag = $null
						$RetentionTagValue = $null
						if ($Item.TryGetProperty($PR_POLICY_TAG,[ref] $RetentionTag))  
						{  
							$RetentionTagValue = $RetentionTag
						} 
						$RetentionTagonItem = $Null
						$RetentionTagonItem = $PolicyTable | ?{$_.PolicyTagByte -eq [String]$RetentionTagValue}
						if(($RetentionTagonItem.PolicyName) -and ($RetentionTagonItem.PolicyName -ne $RetentionTagonFolder.PolicyName))
						{$td = " bgcolor=""#B0C4DE"""} else {$td = $null}
						#HTML Report
						if ($HTMLReport)
						{
							if (($Item.DateTimeReceived) -and ($Item.Size) -and ($RetentionTagValue) -and ($RetentionDateValue))
							{
							$HTMLTable = "<tr bgcolor=$($color)><td width=""20%"">$($fpath)</td><td width=""50%"">$($Item.Subject)</td><td width=""10%"">$($Item.DateTimeReceived)</td><td width=""10%"">$($RetentionDateValue)</td><td width=""5%"">$($Item.Size)</td><td width=""5%"" $($td)>$($RetentionTagonItem.PolicyName)</td></tr>"
							}
							elseif (($Item.DateTimeReceived) -and ($Item.Size) -and (!$RetentionDateValue) -and (!$RetentionTagonItem.PolicyName))
							{
								if(!$HideNeverExpires)
								{
								$HTMLTable = "<tr bgcolor=$($color)><td width=""20%"">$($fpath)</td><td width=""50%"">$($Item.Subject)</td><td width=""10%"">$($Item.DateTimeReceived)</td><td width=""10%"">Never Expires</td><td width=""5%"">$($Item.Size)</td><td width=""5%"" $($td)>No Retention Tag</td></tr>"
								}
							}
							elseif (($Item.DateTimeReceived) -and ($Item.Size) -and (!$RetentionDateValue))
							{
								if(!$HideNeverExpires)
								{
								$HTMLTable = "<tr bgcolor=$($color)><td width=""20%"">$($fpath)</td><td width=""50%"">$($Item.Subject)</td><td width=""10%"">$($Item.DateTimeReceived)</td><td width=""10%"">Never Expires</td><td width=""5%"">$($Item.Size)</td><td width=""5%"" $($td)>$($RetentionTagonItem.PolicyName)</td></tr>"
								}
							}
							Add-Content -Encoding UTF8 $HTMLReportFile $HTMLTable
							$HTMLTable = $Null
						}
						#CSV Report
						if ($CSVReport)
						{
							if (($Item.DateTimeReceived) -and ($Item.Size) -and ($RetentionTagValue) -and ($RetentionDateValue))
							{
							$CSVTable = "$($fpath),""$($Item.Subject)"",$($Item.DateTimeReceived),$($RetentionDateValue),$($Item.Size),$($RetentionTagonItem.PolicyName),"
							}
							elseif (($Item.DateTimeReceived) -and ($Item.Size) -and (!$RetentionDateValue) -and (!$RetentionTagonItem.PolicyName))
							{
								if(!$HideNeverExpires)
								{
									$CSVTable = "$($fpath),""$($Item.Subject)"",$($Item.DateTimeReceived),Never Expires,$($Item.Size),No Retention Tag,"
								}
							}
							elseif (($Item.DateTimeReceived) -and ($Item.Size) -and (!$RetentionDateValue))
							{
								if(!$HideNeverExpires)
								{
								$CSVTable = "$($fpath),""$($Item.Subject)"",$($Item.DateTimeReceived),Never Expires,$($Item.Size),$($RetentionTagonItem.PolicyName),"
								}
							}
							Add-Content -Encoding UTF8 $CSVReportFile $CSVTable
							$CSVTable = $Null
						}
				}
				#The Do loop will handle any paging that is required if there are more the 1000 items in the folder  
				$ivItemView.offset += $fiItemResult.Items.Count
			}
			while($fiItemResult.MoreAvailable -eq $true)
			$ivItemView.offset = 0
		}
	}

#HTML Footer
if ($HTMLReport)
{
	$HTMLFooter = "</table></html></body>"
	#Finish HTML Report
	Add-Content -Encoding UTF8 $HTMLReportFile $HTMLFooter
}

$EndTime = Get-Date
$ElapsedTime = $EndTime-$StartTime
"##########################################################################################"
"Elapsed Time: " + $ElapsedTime.Minutes + ":" + $ElapsedTime.Seconds
if ($CSVReport)
{
$Path = Get-Item $CSVReportFile
"##########################################################################################"
"CSV Report: " + $Path.versioninfo.filename
"##########################################################################################"
}
if ($HTMLReport)
{
$Path = Get-Item $HTMLReportFile
"##########################################################################################"
"HTML Report: " + $Path.versioninfo.filename
"##########################################################################################"
}
#------------------------------------------------------------------------------  
#  
# Copyright © 2015 Microsoft Corporation.  All rights reserved.  
#  
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT  
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT  
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS  
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR   
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.  
#  
#------------------------------------------------------------------------------  
#  
# PowerShell Source Code  
#  
# NAME:  
#    Azure_File_Service.ps1  
#  
# VERSION:  
#    1.3 
#  
#------------------------------------------------------------------------------  
 

"------------------------------------------------------------------------------ " | Write-Host -ForegroundColor Yellow 
""  | Write-Host -ForegroundColor Yellow 
" Copyright © 2015 Microsoft Corporation.  All rights reserved. " | Write-Host -ForegroundColor Yellow 
""  | Write-Host -ForegroundColor Yellow 
" THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED `“AS IS`” WITHOUT " | Write-Host -ForegroundColor Yellow 
" WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT " | Write-Host -ForegroundColor Yellow 
" LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS " | Write-Host -ForegroundColor Yellow 
" FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR  " | Write-Host -ForegroundColor Yellow 
" RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. " | Write-Host -ForegroundColor Yellow 
"------------------------------------------------------------------------------ " | Write-Host -ForegroundColor Yellow 
""  | Write-Host -ForegroundColor Yellow 
" PowerShell Source Code " | Write-Host -ForegroundColor Yellow 
""  | Write-Host -ForegroundColor Yellow 
" NAME: " | Write-Host -ForegroundColor Yellow 
"    Azure_File_Service.ps1 " | Write-Host -ForegroundColor Yellow 
"" | Write-Host -ForegroundColor Yellow 
" VERSION: " | Write-Host -ForegroundColor Yellow 
"    1.3" | Write-Host -ForegroundColor Yellow 
""  | Write-Host -ForegroundColor Yellow 
"------------------------------------------------------------------------------ " | Write-Host -ForegroundColor Yellow 
"" | Write-Host -ForegroundColor Yellow 
"`n This script SAMPLE is provided and intended only to act as a SAMPLE ONLY," | Write-Host -ForegroundColor Yellow 
" and is NOT intended to serve as a solution to any known technical issue."  | Write-Host -ForegroundColor Yellow 
"`n By executing this SAMPLE AS-IS, you agree to assume all risks and responsibility associated."  | Write-Host -ForegroundColor Yellow 

$ContinueAnswer = Read-Host "`n`tDo you wish to proceed at your own risk? (Y/N)" 
If ($ContinueAnswer -ne "Y") { Write-Host "`n Exiting." -ForegroundColor Red;Exit }

Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

#import the Azure PowerShell module
Write-Host "`n[WORKITEM] - Importing Azure PowerShell module" -ForegroundColor Yellow

If ($ENV:Processor_Architecture -eq "x86")
{
        $ModulePath = "$Env:ProgramFiles\Microsoft SDKs\Azure\PowerShell\ServiceManagement\Azure\Azure.psd1"

}
Else
{
        $ModulePath = "${env:ProgramFiles(x86)}\Microsoft SDKs\Azure\PowerShell\ServiceManagement\Azure\Azure.psd1"
}

Try
{
        If (-not(Get-Module -name "Azure")) 
        { 
               If (Test-Path $ModulePath) 
               { 
                       Import-Module -Name $ModulePath
               }
               Else
               {
                       #show module not found interaction and bail out
                       Write-Host "[ERROR] - Azure PowerShell module not found. Exiting." -ForegroundColor Red
                       Exit
               }
        }

        Write-Host "`tSuccess"
}
Catch [Exception]
{
        #show module not found interaction and bail out
        Write-Host "[ERROR] - PowerShell module not found. Exiting." -ForegroundColor Red
        Exit
}

#Check the Azure PowerShell module version
Write-Host "`n[WORKITEM] - Checking Azure PowerShell module verion" -ForegroundColor Yellow
$APSMajor =(Get-Module azure).version.Major
$APSMinor =(Get-Module azure).version.Minor
$APSBuild =(Get-Module azure).version.Build
$APSVersion =("$PSMajor.$PSMinor.$PSBuild")

If ($APSVersion -ge 0.8.16)
{
    Write-Host "`tSuccess"
}
Else
{
   Write-Host "[ERROR] - Azure PowerShell module must be version 0.8.16 or higher. Exiting." -ForegroundColor Red
   Exit
}

#Use Add-AzureAccount
Write-Host "`n[INFO] - Authenticating Azure account."  -ForegroundColor Yellow
Add-AzureAccount | out-null

#Check to make sure authentication occured
If ($?)
{
	Write-Host "`tSuccess"
}
Else
{
	Write-Host "`tFailed authentication" -ForegroundColor Red
	Exit
}

#####
#Azure subscription selection
#####
Write-Host "`n[INFO] - Obtaining subscriptions" -ForegroundColor Yellow
[array] $AllSubs = Get-AzureSubscription 

If ($AllSubs)
{
        Write-Host "`tSuccess"

        #$AllSubs | FL 
}
Else
{
        Write-Host "`tNo subscriptions found. Exiting." -ForegroundColor Red
        "`tNo subscriptions found. Exiting." 
        Exit
}

Write-Host "`n[SELECTION] - Select the Azure subscription." -ForegroundColor Yellow

$SelSubName = $AllSubs | Out-GridView -PassThru -Title "Select the Azure subscription"

If ($SelSubName)
{
	#Write sub
	Write-Host "`tSelection: $($SelSubName.SubscriptionName)"
		
        $SelSub = $SelSubName.SubscriptionId
        Select-AzureSubscription -SubscriptionId $SelSub | Out-Null
}
Else
{
        Write-Host "`n[ERROR] - No Azure subscription was selected. Exiting." -ForegroundColor Red
        Exit
}

Write-Host "`n[SELECTION] - Input for script workload." -ForegroundColor Yellow

$input0 = new-object psobject
Add-Member -InputObject $input0 -MemberType NoteProperty -Name Workload -Value "Create Azure Storage Account" -Force
$input1 = new-object psobject
Add-Member -InputObject $input1 -MemberType NoteProperty -Name Workload -Value "Create a Azure File Share" -Force
$input2 = new-object psobject
Add-Member -InputObject $input2 -MemberType NoteProperty -Name Workload -Value "Remove a Azure File Share" -Force
$input3 = new-object psobject
Add-Member -InputObject $input3 -MemberType NoteProperty -Name Workload -Value "Print NET USE command for Windows IaaS VM" -Force
$input4 = new-object psobject
Add-Member -InputObject $input4 -MemberType NoteProperty -Name Workload -Value "Print Sudo Mount command for Linux IaaS VM" -Force
$input5 = new-object psobject
Add-Member -InputObject $input5 -MemberType NoteProperty -Name Workload -Value "Create a directory in Azure Files Share" -Force
Add-Member -InputObject $input5 -MemberType NoteProperty -Name Info -Value "Allows for the creation of a folder off the root of the Azure File share" -Force
$input6 = new-object psobject
Add-Member -InputObject $input6 -MemberType NoteProperty -Name Workload -Value "Create one gigabyte test file" -Force
Add-Member -InputObject $input6 -MemberType NoteProperty -Name Info -Value "One gigabyte test transfer file" -Force
$input7 = new-object psobject
Add-Member -InputObject $input7 -MemberType NoteProperty -Name Workload -Value "Upload a local file to Azure Files" -Force
Add-Member -InputObject $input7 -MemberType NoteProperty -Name Info -Value "Allows for the upload of a single file to a Azure File Share Directory" -Force
$input8 = new-object psobject
Add-Member -InputObject $input8 -MemberType NoteProperty -Name Workload -Value "Download a file from Azure Storage" -Force
Add-Member -InputObject $input8 -MemberType NoteProperty -Name Info -Value "Allows for the download of a single file from Azure File Share Directory" -Force

[array] $Input += $input0
[array] $Input += $input1
[array] $Input += $input2
[array] $Input += $input3
[array] $Input += $input4
[array] $Input += $input5
[array] $Input += $input6
[array] $Input += $input7
[array] $Input += $input8

$Work = $Input | Select-Object Workload,Info | Out-GridView  -Title "Select workload for script" -PassThru
$SelWork = $Work.Workload

Write-Host "`n[WORKITEM] - Script will attempt to $($SelWork.tolower())" -ForegroundColor Yellow

if ($SelWork -eq "Create Azure Storage Account")
{
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')  
  	$StorName = [Microsoft.VisualBasic.Interaction]::InputBox("Create a name for Azure storage account", "Storage Account", "storage")  
  	$StorName = $StorName.tolower()
	$SpecChars = '!', '"', '£', '$', '%', '&', '^', '*', '(', ')', '@', '=', '+', '¬', '`', '\', '<', '>', '.', '?', '/', ':', ';', '#', '~', "'", '-', '_', ' '
	$RemSpecChars = [string]::join('|', ($SpecChars | % {[regex]::escape($_)}))
	$StorName = $StorName -replace $RemSpecChars, ""
	Write-Host "`n[SELECTION] - Select the Azure VM Location." -ForegroundColor Yellow 
	$SelLocationName = Get-AzureLocation  
	$GEOselection = $SelLocationName | select DisplayName | Sort-Object DisplayName | Out-GridView -Title "Select Region" -passthru 
	$Loc = $SelLocationName | Where {($_.DisplayName -eq $GEOselection.DisplayName)} 
	$region = $Loc.Name 
	$StorAccName = "storage$StorName" 
	 
	New-AzureStorageAccount -StorageAccountName $StorAccName -Label $StorName -Location $region -Description "Storage Account for $StorName" -Type Standard_GRS -WarningAction SilentlyContinue | out-null 
	 
	#Check to make sure AzureStorageAccount was created 
	$CreatedStorageAccount = Get-AzureStorageAccount -StorageAccountName $StorAccName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue 
	 
	If ($CreatedStorageAccount) 
	{ 
	    Write-Host "`n[SUCCESS] Script is created storage account in the region $region" -ForegroundColor Green
	} 
	Else 
	{ 
	    Write-Host "`tFailed to create Storage Account" -ForegroundColor Red 
	    Exit 
	} 
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Create a Azure File Share")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')  
  	$sharename = [Microsoft.VisualBasic.Interaction]::InputBox("Input name for Azure File Share", "Share Name", "azureshare")  
  	$sharename=$sharename.tolower()
	$SpecChars = '!', '"', '£', '$', '%', '&', '^', '*', '(', ')', '@', '=', '+', '¬', '`', '\', '<', '>', '.', '?', '/', ':', ';', '#', '~', "'", '-', '_', ' '
	$RemSpecChars = [string]::join('|', ($SpecChars | % {[regex]::escape($_)}))
	$sharename = $sharename -replace $RemSpecChars, ""
	Try
	{
    	$share = New-AzureStorageShare -Name $sharename -Context $ctx
	}
	Catch [Exception]
	{
		Write-Host $_ -ForegroundColor Red
		Exit
	}
	
	If (!($share))
	{
		$ErrorMessage = $Error[0].ToString()
		Write-Host "`n[ERROR] $ErrorMessage`n Exiting`n`n" -ForegroundColor Red
		exit
	}

    $Name = $NULL
    $Name = (Get-AzureStorageShare -Context $ctx -Name $sharename).Name
        If ($Name)
        {
                 Write-Host "`n[INFO] Azure file share $sharename created`n" -ForegroundColor Green
        }
        Else
        {
               Write-Host "`n[ERROR] Azure file share $sharename failed`n" -ForegroundColor Red
        }
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Remove a Azure File Share")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Storage Share to Remove" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	
	Remove-AzureStorageShare -Name $SelSourceShareName -Context $ctx -confirm:$false 

        If ($SelSourceShareName)
        {
               Write-Host "`n Azure file share $SelSourceShareName removal succeeded`n" -ForegroundColor Green
        }
        Else
        {
               Write-Host "`n Azure file share $SelSourceShareName removal failed`n" -ForegroundColor Red
        }
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Print NET USE command for Windows IaaS VM")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Azure File Share" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	if ($SelSourceShareName -eq $null) {exit}
	Else
	{
	Write-host "`n The command is:"
	Write-host "`n net use z: \\$($StorageAccName).file.core.windows.net\$($SelSourceShareName) /u:$($StorageAccName) $Key" -ForegroundColor Green
	Write-host "`n The SMB share can be accessed from any Azure node (VM/Worker/Web role) hosted in the same region as the storage account hosting the share."
	Write-host "`n"
	}
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Print Sudo Mount command for Linux IaaS VM")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Azure File Share" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	if ($SelSourceShareName -eq $null) {exit}
	Else
	{
	Write-host "`n The command is:"
	Write-host "`n sudo mount -t cifs //$($StorageAccName).file.core.windows.net/$($SelSourceShareName) ./mymountpoint -o vers=2.1,username=$($StorageAccName),password=$($Key),dir_mode=0777,file_mode=0777" -ForegroundColor Green
	Write-host "`n The SMB share can be accessed from any Azure node (VM/Worker/Web role) hosted in the same region as the storage account hosting the share."
	Write-host "`n"
	}
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Create a directory in Azure Files Share")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Azure File Share" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	Set-AzureSubscription -SubscriptionId $SelSub -CurrentStorageAccountName $StorageAccName | Out-Null
	
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')  
  	$DirName = [Microsoft.VisualBasic.Interaction]::InputBox("Create a Directory for Azure File Share", "Directory Name", "directory")  
  	$DirName=$DirName.tolower()
	$SpecChars = '!', '"', '£', '$', '%', '&', '^', '*', '(', ')', '@', '=', '+', '¬', '`', '\', '<', '>', '.', '?', '/', ':', ';', '#', '~', "'", '-', '_', ' '
	$RemSpecChars = [string]::join('|', ($SpecChars | % {[regex]::escape($_)}))
	$DirName = $DirName -replace $RemSpecChars, ""
	Try
	{
    	New-AzureStorageDirectory -ShareName $SelSourceShareName -Path $DirName
	}
	Catch [Exception]
	{
		Write-Host $_ -ForegroundColor Red
		Exit
	}
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Upload a local file to Azure Files")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Azure File Share" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	Set-AzureSubscription -SubscriptionId $SelSub -CurrentStorageAccountName $StorageAccName | Out-Null
	$AzureFileDir = Get-AzureStorageFile -ShareName $SelSourceShareName
	If ($AzureFileDir -eq $null) {Write-Host "`n[ERROR] No Azure File Directory found." -ForegroundColor Red;Exit}
	$SelAzureFileDir = $AzureFileDir | Out-GridView -Title "Select Azure File Share Directory" -PassThru

	#Pick file to upload
	$upload = Get-FileName -initialDirectory "c:\"
	
	Try
	{
    	Set-AzureStorageFileContent -ShareName $SelSourceShareName -Source $upload -Path $SelAzureFileDir.Name
		Write-Host "`n[SUCCESS] file $($upload) uploaded to Azure Files" -ForegroundColor Green
	}
	Catch [Exception]
	{
		Write-Host $_ -ForegroundColor Red
		Exit
	}
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Download a file from Azure Storage")
{
	$AllStorageAccounts = Get-AzureStorageAccount -WarningAction SilentlyContinue | where {$_.Endpoints -match "file.core.windows.net"} 
	If ($AllStorageAccounts -eq $null) {Write-Host "`n[ERROR] No Azure File Endpoints found." -ForegroundColor Red;Exit}
	$SelStorageAccount = $AllStorageAccounts | Select-Object Label,StorageAccountName,Location,AffinityGroup,AccountType | Out-GridView -Title "Select Storage Account" -PassThru
	$StorageAccName = $SelStorageAccount.StorageAccountName
    $Key = (Get-AzureStorageKey -StorageAccountName $StorageAccName).Primary
    $ctx = New-AzureStorageContext -StorageAccountName $StorageAccName -StorageAccountKey $key
	$SelAzureStorageShare = Get-AzureStorageShare -Context $ctx | Out-GridView -Title "Select Azure File Share" -PassThru
	$SelSourceShareName = $SelAzureStorageShare.Name
	Set-AzureSubscription -SubscriptionId $SelSub -CurrentStorageAccountName $StorageAccName | Out-Null
	$AzureFileDir = Get-AzureStorageFile -ShareName $SelSourceShareName
	If ($AzureFileDir -eq $null) {Write-Host "`n[ERROR] No Azure File Directory found." -ForegroundColor Red;Exit}
	$SelAzureFileDir = $AzureFileDir | Out-GridView -Title "Select Azure File Share Directory" -PassThru
	$AzureFile = Get-AzureStorageFile -ShareName $SelSourceShareName -Path $SelAzureFileDir.Name | Out-GridView -Title "Select Azure File Share item to Download" -PassThru
	$SelAzureFileDirName = $SelAzureFileDir.Name
	$AzureFileName = $AzureFile.Name
	
	Add-Type -AssemblyName System.Windows.Forms
	$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$FolderBrowser.ShowDialog()
	$FolderBrowser.SelectedPath | Out-Null
	$Download = $FolderBrowser.SelectedPath
	
		Try
	{
    	Get-AzureStorageFileContent -ShareName $SelSourceShareName -Path $SelAzureFileDirName'/'$AzureFileName -Destination $Download
		Write-Host "`n[SUCCESS] file $($AzureFileName) downloaded to $($Download)" -ForegroundColor Green
	}
	Catch [Exception]
	{
		Write-Host $_ -ForegroundColor Red
		Exit
	}
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit
}
If ($SelWork -eq "Create one gigabyte test file")
{
	Add-Type -AssemblyName System.Windows.Forms
	$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$FolderBrowser.ShowDialog()
	$FolderBrowser.SelectedPath | Out-Null
	$SelFol = $FolderBrowser.SelectedPath
	fsutil file createnew $SelFol'\'1Gbfile.txt 1073741824
	
Write-Host "`n Press any key to continue ...`n"
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Exit	
}


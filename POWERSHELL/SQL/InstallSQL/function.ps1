#. ./common.ps1


function displayDrives  {
Get-WmiObject -class "Win32_LogicalDisk" | ?{ @(2, 3) -contains $_.DriveType } | where {$_.Freespace} | select Name, VolumeName, Size, FreeSpace
}

function Get-instacneName
{
    $InstanceName=(Get-Content .\variables.txt)[3]
    return $InstanceName
}


function Get-serverenvironment
{
    $serverEnvironment=(Get-Content .\variables.txt)[8]
    return $serverEnvironment
}

function get-SQlserverversion
{
    $srclocation=(Get-Content .\variables.txt)[0]
    $regpattern = '\d\d\d\d'
    $vers = $srclocation | Select-String $regpattern -AllMatches
    return $vers.Matches.Value | unique
}

function get-SQlservermedia
{
	$SQlservermedialocation=(Get-Content .\variables.txt)[0]
    return $SQlservermedialocation
}

function get-SQlserverpatchmedia
{
	$SQlserverPatchSourcelocation=(Get-Content .\variables.txt)[1]
    return $SQlserverPatchSourcelocation
}

function Get-serviceaccount
{
	$SQLServerServiceAccountName=(Get-Content .\variables.txt)[2]
    return $SQLServerServiceAccountName
}

function get-Serviceaccountpassword
{
    [System.Security.SecureString]$SvcPasswordSec = Read-Host "Enter the Service Account password: " -AsSecureString; 
    [String]$SvcPasswordSec1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SvcPasswordSec))
    #[System.Security.SecureString]$saPasswordSec = Read-Host "Enter the sa password: " -AsSecureString; 
    #[String]$SQLServerServiceaccountpassword1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SQLServerServiceaccountpassword))
    return $SvcPasswordSec1
}

function get-sapassword
{
    [System.Security.SecureString]$SPasswordSec = Read-Host "Enter the SA password: " -AsSecureString; 
    [String]$SPasswordSec1 = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SPasswordSec))
    return $SPasswordSec1
}

function Datadriveletter
{
	$DatDriveLetter=(Get-Content .\variables.txt)[4]
    return $DatDriveLetter
}

function logdriveletter
{
	$LogDriveLetter=(Get-Content .\variables.txt)[5]
    return $LogDriveLetter
}

function SSISDriveLetter
{
	$SSISDriveLetter=(Get-Content .\variables.txt)[6]
    return $SSISDriveLetter
}

function joboutputdriveletter
{
	$jobopDriveLetter=(Get-Content .\variables.txt)[7]
    return $jobopDriveLetter
}

function tempdbdriveletter
{
	$tempdbdriveletter=(Get-Content .\variables.txt)[9]
    return $tempdbdriveletter
}

function SQLCUmedia
{
    $SQLserverCUloc=(Get-Content .\variables.txt)[10]
    return $SQLserverCUloc
}

function Update-SQLBackupFolderPermissions{
    $ServerName="SACODBA03"
    #$serverEnvironment=Get-serverenvironment
    $serverEn_bkp=Get-serverenvironment
    if ($serverEn_bkp -like "Dev*"){$serverEnvironment="Development"}
    if ($serverEn_bkp -like "Test*"){$serverEnvironment="Test"}
    if ($serverEn_bkp -like "Prod*"){$serverEnvironment="Production"}
    write-host "server environment is: " $serverEnvironment
    $BackupRoot = "D:\Backups\"+$serverEnvironment
    $SQLServerServiceAccount = Get-serviceaccount
    Write-Host "SQL Server Service is running as:$SQLServerServiceAccount"
    Write-Host "Backup root as: $BackupRoot"

    $session = New-PSSession -ComputerName $ServerName;
    Invoke-Command -Session $session -Args $BackupRoot, $SQLServerServiceAccount -ScriptBlock    {
        param([string]$BackupRoot,[string]$SQLServerServiceAccount)
        $acl = Get-Acl $BackupRoot;
        $accessrule = New-Object System.Security.AccessControl.FileSystemAccessRule($SQLServerServiceAccount,"FullControl","ContainerInherit,ObjectInherit","None","Allow")
        $acl.AddAccessRule($accessRule)
        Set-Acl -aclobject $acl $BackupRoot
    };
    Remove-PSSession $session;
}

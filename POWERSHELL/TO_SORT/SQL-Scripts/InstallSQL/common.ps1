
function displayDrives {

Get-WmiObject -class "Win32_LogicalDisk" | ?{ @(2, 3) -contains $_.DriveType } | where {$_.Freespace} | select Name, VolumeName, Size, FreeSpace
}

if (Test-Path ".\variables.txt")
{
  Remove-Item ".\variables.txt"
}


$getSQlservermedia = read-host "Enter media location"
$getSQlserverpatchmedia = read-host "Enter patch location with .EXE file name"
$Getserviceaccount = read-host "Enter service account with PMMR\account name"
$GetinstacneName = read-host "Enter InstanceName"
$Datadriveletter = read-host "Enter data drive letter"

if(!(Test-Path ($Datadriveletter+':')))
{
    write-host " Data drive" $Datadriveletter": does not exist, here is the disk layout" -foregroundcolor red -backgroundcolor yellow
    displayDrives
    exit
}
$logdriveletter = read-host "Enter log drive letter"

if(!(Test-Path ($LogDriveLetter+':')))
{
    write-host " Log drive" $LogDriveLetter": does not exist, here is the disk layout" -foregroundcolor red -backgroundcolor yellow
    displayDrives
    exit
}

$SSISdriveletter = read-host "Enter SSIS drvie letter"
$jobopletter = read-host "Enter joboutput drvie letter" 
$serverenvironment = read-host "Enter Server environment like test,test02 etc.." 

if(!(Test-Path ('T:')))
{
    $tempdbdriveletter = read-host "T: drive does not exist.  Please enter TempDB data drvie letter" 
    if(!(Test-Path ($tempdbdriveletter+':')) -or (Test-Path ('T:')))
    {
        write-host " Tempdbdriveletter or " $tempdbdriveletter": or dedicated drive T does not exist, here is the disk layout" -foregroundcolor red -backgroundcolor yellow
        displayDrives 
        exit
    }
}
else
{
    $tempdbdriveletter = 'T'
}

$yn = Read-Host "Do we need to install SQL server Cummulative Update ? Yes or No"
If($yn -like "y*")
{
    $SQLserverCUmedialocation = Read-host "Enter SQLserver CU media location"
}

else
{
    Write-Host "SQL server CU not required"
    $SQLserverCUmedialocation = 'null'

}



write-output $getSQlservermedia $getSQlserverpatchmedia $Getserviceaccount $GetinstacneName $Datadriveletter $logdriveletter $SSISdriveletter $jobopletter $serverenvironment $tempdbdriveletter $SQLserverCUmedialocation >>".\variables.txt"

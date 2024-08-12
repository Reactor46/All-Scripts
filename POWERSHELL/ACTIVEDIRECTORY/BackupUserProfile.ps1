<#
.SYNOPSIS
    Performs a backup of profiles based usernames defined in a text file and, optionally, delete profile

.DESCRIPTION
    This script performs a backup of user profile to a remote server, log the information on log file and
    delete the profile in case of specification. Before delete, it compares if source and destination
    have the same size and delete only if it is equal. If you don't want that, change the script. The
    SourceUserFile is a text file with usernames of profile to backup and, eventually, delete. Use one
    user profile per line.

.AUTHOR
    Gustavo Zimmermann Montesdioca - Gustavo_Percio@hotmail.com

.DATE
    Nov 02, 2015 

.PARAMETER Source
    Mandatory: Path to profile folder (i.e. C:\Users)

.PARAMETER SourceUserFile
    Mandatory: The text file with usernames of profiles to backup (one username per line)

.PARAMETER Destination
    Mandatory: The server share (UNC) you want to backup user profiles (ex: \\SERVER\SHARE)

.PARAMETER DeleteSourceProfile
    Mandatory: Delete Source Profile (yes or no). Default: no.

#>
#$Source='C:\Users'
#$SourceUserFile='c:\users\administrator\desktop\backup.txt'
#$Destination='\\gm9-cus-db03\backup'
#$DeleteSourceProfile="yes"

#Get parameters required for this script to run
Param
(

[Parameter(Mandatory=$true)]
        [String] $Source,
[Parameter(Mandatory=$true)]
        [String] $SourceUserFile,
[Parameter(Mandatory=$true)]
        [String] $Destination,
[Parameter(Mandatory=$true)]
        [String] $DeleteSourceProfile='no'

)

$validation=Test-Path $Destination
$SourceUsernames = Get-Content $SourceUserfile
$BackupFolder=Get-Date -UFormat "%Y%m%d-%H%M%S-ProfileBackup" #The backup folder in formation YYYYMMDD-HHMMSS-profilebackup

New-PSDrive -Name "ProfileBackup" -PSProvider FileSystem -Root $Destination -ErrorAction 'silentlycontinue'

if ($validation -eq $True)
{      
        Set-Location ProfileBackup:
}
else
{
    Write-Host "Error! The destination is unreachable"
    break
}

Foreach($uid in $SourceUsernames)
{
    if($uid -ne "")
    {
        Write-Host ""
        Write-Host "Backuping profile $uid"
        robocopy $Source\$uid $Destination\$BackupFolder\$uid *.* /s /e /sec /mir /R:1 /W:5 /log:$BackupFolder-$uid'.log'
            
        $SourceSize=Get-ChildItem -Path "$Source\$uid\*.*" -Recurse -Force -ErrorAction 'silentlycontinue' | Measure-Object -property length -sum
        $DestSize=Get-ChildItem -Path "$Destination\$BackupFolder\$uid\*.*" -Recurse -Force -ErrorAction 'silentlycontinue' | Measure-Object -property length -sum
            
         
        if($DeleteSourceProfile -eq "yes")
        {
            if([int]$SourceSize.Sum -eq [int]$DestSize.Sum)
            {
                Write-Host ""
                Write-Host "Deleting profile $uid"
                mkdir $Source\empty_dir
                robocopy empty_dir $Source\$uid /s /mir /log:$BackupFolder-$uid'-Removal.log'
                rmdir $Source\$uid
                rmdir $Source\empty_dir
            }
            else
            {
                Write-Host "It did not delete profile $uid due backup is not consistent"
            }
        }
    }
}

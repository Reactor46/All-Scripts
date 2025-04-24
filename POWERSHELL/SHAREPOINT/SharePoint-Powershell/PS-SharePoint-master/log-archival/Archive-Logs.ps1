<#
.SYNOPSIS
  Archive log files

.DESCRIPTION
  This script can move and compress files from specified location with given retention period. 
  It has been created to archive log files on regular basis.

.PARAMETER SourceFolder
  The folder where the files to be archived are stored

.PARAMETER TargetRootFolder
  Target location for the archived files. Running account must have write permissions to this folder.
  If not provided, the source location will be used to store the zip archive.

.PARAMETER TempFolderRoot
  Temporary location for the archival folder. Running account must have write permissions to this folder.
  Default value: user profile temp folder

.PARAMETER RetentionDays
  Number of days to keep the files

.PARAMETER KeepTempFiles
  Set this on to keep temporary files after archival

.PARAMETER Delete
  Makes the script to delete files instead of moving them to target location. If you set this, there is no need to set temporary and target folder locations.

.INPUTS
  None

.OUTPUTS
  Zip archive with files from source folder. 
  File name is <Last modified date of the oldest file>-<Last modified date of the newest file>.zip

.NOTES
	Author: Madhu Muppavarapu, Egor Yudkin
	Version: 2.1
	|
	Release history:
	v2.1 [04/13/2018]
		Added an option to delete files instead of moving with "Delete" switch
	v2.0 [04/03/2018]
		Renamed some arguments [Breaking change]
		Added the target folder argument
		Documentation
		Added default temp folder location
		Changed default behavior to clean up temp location
	v1.1 [03/05/2018]
		Refactoring
	v1.0 Initial script development

   
.EXAMPLE
  Archive IIS logs for Deafult Web Site which are older than 30 days from default location to "e:\LogsArchive\IIS\default" folder.
  User temp folder will be used for temporary files, which will be deleted after script is completed.
  
  .\Archive-Logs.ps1 -SourceFolder "c:\intepub\logs\w3svc1" -TargetRootFolder "e:\LogsArchive\IIS\default" -RetentionDays 30

.EXAMPLE
  Delete IIS logs for Deafult Web Site which are older than 30 days from default location to "e:\LogsArchive\IIS\default" folder.
  
  .\Archive-Logs.ps1 -SourceFolder "c:\intepub\logs\w3svc1" -Delete -RetentionDays 30
#>

#region [Script Parameters]
#---------------------------------------------------------------------------------------------------------------
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [ValidateScript({
        if ((Test-Path -PathType Container $_) -ne $true)
        {
            throw "The SourceFolder parameter is not correct. Please provide existing folder"
        }
        else 
        {
            return $true
        } 
    })]
    [string]
    $SourceFolder,

    [parameter()]
    [parameter(ParameterSetName = "MoveFiles")]
    [ValidateScript({
    if ((Test-Path -PathType Container $_) -ne $true)
        {
            $_ = New-Item -ItemType directory -Path $_
            if ((Test-Path -PathType Container $_) -ne $true) 
            {
                throw "The TargetRootFolder parameter is not correct. Please provide proper folder name"
            }
            else 
            {
                return $true
            }
        }
        else 
        {
            return $true
        } 
    })]
    [string]
    $TargetRootFolder,

    [parameter()]
    [parameter(ParameterSetName = "MoveFiles")]
    [ValidateScript({
    if ((Test-Path -PathType Container $_) -ne $true)
        {
            $_ = New-Item -ItemType directory -Path $_
            if ((Test-Path -PathType Container $_) -ne $true) 
            {
                throw "The TempFolderRoot parameter is not correct. Please provide proper folder name"
            }
            else 
            {
                return $true
            }
        }
        else 
        {
            return $true
        } 
    })]
    [string]
    $TempFolderRoot = $env:temp,

    [parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [uint16]
    $RetentionDays,

    [parameter()]
    [parameter(ParameterSetName = "MoveFiles")]
    [switch]
    $KeepTempFiles,

    [parameter()]
    [parameter(ParameterSetName = "DeleteFiles")]
    [switch]
    $Delete
)
#endregion

# Calculate retention date
$currentDate = Get-Date
$date = $currentDate.ToString("yyyy-MM-dd-hh-mm")
$retentionDate = $currentDate.AddDays(-$RetentionDays)

# get the list of files in the original folder which are older than $retentionDays and move the files to a temporary location
$items = Get-ChildItem -Path $SourceFolder | Where-Object {$_.LastWriteTime -lt $retentionDate}

# calculate a date range for impacted files
$oldestDate = $items | Sort-Object LastWriteTime | Select -First 1 -Property LastWriteTime
$newestDate = $items | Where-Object {$_.LastWriteTime -lt $retentionDate} | Sort-Object LastWriteTime | Select -Last 1 -Property LastWriteTime

# move files
if ($items.Count -gt 0)
{
    if ($Delete)
    {
        $items | Remove-Item
        return
    }
    else
    {
        $tempFinalFolder = New-Item -ItemType Directory -Path (Join-Path $TempFolderRoot $date) -Force
        $items | Move-Item -destination $tempFinalFolder
    }
}
else
{
    Write-Output "Nothing to archive"
    return
}

$CompressionToUse = [System.IO.Compression.CompressionLevel]::Optimal
$IncludeBaseFolder = $false
if ($TargetRootFolder)
{
    $zipTo = "{0}\{1}-{2}.zip" -f $TargetRootFolder,$oldestDate.LastWriteTime.ToString("yyyy-MM-dd"),$newestDate.LastWriteTime.ToString("yyyy-MM-dd")
}
else
{
    $zipTo = "{0}\{1}-{2}.zip" -f $SourceFolder,$oldestDate.LastWriteTime.ToString("yyyy-MM-dd"),$newestDate.LastWriteTime.ToString("yyyy-MM-dd")
}

#add the files in the temporary location to a zip folder
[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempFinalFolder, $ZipTo, $CompressionToUse, $IncludeBaseFolder)

if (!$KeepTempFolder)
{
    Remove-Item $tempFinalFolder -Recurse -Force
}
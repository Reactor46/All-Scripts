##
# Import website (IIS) logs into SQL Server using Log Parser and PowerShell
# http://www.technologytoolbox.com/blog/jjameson/archive/2012/02/29/import-website-iis-logs-into-sql-server-using-log-parser.aspx
#
## This script will...
#
#     1. Extract (a.k.a. unzip) the log files in the /httplog folder (which I
#        periodically FTP from the Production environment) and subsequently move
#        the zip files to the /httplog/archive folder.
#     2. Import the log files using the LogParser utility.
#     3. Remove the log files from the /httplog folder (to avoid inserting
#        duplicate data the next time the script is run).
#
## To create the SQL database table schema:
#
#     CREATE TABLE dbo.WebsiteLog
#     (
#         LogFilename VARCHAR(255) NOT NULL,
#         RowNumber INT NOT NULL,
#         EntryTime DATETIME NOT NULL,
#         SiteName VARCHAR(255) NOT NULL,
#         ServerName VARCHAR(255) NOT NULL,
#         ServerIpAddress VARCHAR(255) NOT NULL,
#         Method VARCHAR(255) NOT NULL,
#         UriStem VARCHAR(255) NOT NULL,
#         UriQuery VARCHAR(255) NULL,
#         Port INT NOT NULL,
#         Username VARCHAR(255) NULL,
#         ClientIpAddress VARCHAR(255) NOT NULL,
#         HttpVersion VARCHAR(255) NOT NULL,
#         UserAgent VARCHAR(255) NOT NULL,
#         Cookie VARCHAR(255) NULL,
#         Referrer VARCHAR(255) NULL,
#         Hostname VARCHAR(255) NOT NULL,
#         HttpStatus INT NOT NULL,
#         HttpSubstatus INT NOT NULL,
#         Win32Status INT NOT NULL,
#         BytesFromServerToClient INT NOT NULL,
#         BytesFromClientToServer INT NOT NULL,
#         TimeTaken INT NOT NULL
#     )
#
## Usage:
#
#     PS C:\> & ".\Import-Website-Log-Files.ps1"
#     Creating archive folder for compressed log files...
#     Extracting compressed log files...
#     Importing log files to database...
#
#     Statistics:
#     -----------
#     Elements processed: 210943
#     Elements output:    210943
#     Execution time:     155.13 seconds (00:02:35.13)
#
#     Removing log files...
#     Successfully imported log files.
#
## Dependencies:
#
#     - Log Parser: https://www.microsoft.com/en-us/download/details.aspx?displaylang=en&id=24659
#
## Settings:
#
#     You'll need to update the log file path and connection string.
#
##

$ErrorActionPreference = "Stop"

Import-Module Pscx -EA 0

function ExtractLogFiles([string] $httpLogPath)
{
    if ([string]::IsNullOrEmpty($httpLogPath) -eq $true)
    {
        throw "The log path must be specified."
    }

    [string] $httpLogArchive = $httpLogPath + "\archive"
    if ((Test-Path $httpLogArchive) -eq $false)
    {
        Write-Host "Creating archive folder for compressed log files..."
        New-Item -ItemType directory -Path $httpLogArchive | Out-Null
    }

    Write-Host "Extracting compressed log files..."
    Get-ChildItem $httpLogPath -Filter "*.zip" |
        ForEach-Object {
            Expand-Archive $_ -OutputPath $httpLogPath
            Move-Item $_.FullName $httpLogArchive
        }
}

function ImportLogFiles([string] $httpLogPath)
{
    if ([string]::IsNullOrEmpty($httpLogPath) -eq $true)
    {
        throw "The log path must be specified."
    }

    [string] $logParser = "${env:ProgramFiles(x86)}" `
        + "\Log Parser 2.2\LogParser.exe"

    [string] $query = `
        [string] $query = `
        "SELECT" `
            + " LogFilename" `
            + ", RowNumber" `
            + ", TO_TIMESTAMP(date, time) AS EntryTime" `
            + ", s-sitename AS SiteName" `
            + ", s-computername AS ServerName" `
            + ", s-ip AS ServerIpAddress" `
            + ", cs-method AS Method" `
            + ", cs-uri-stem AS UriStem" `
            + ", cs-uri-query AS UriQuery" `
            + ", s-port AS Port" `
            + ", cs-username AS Username" `
            + ", c-ip AS ClientIpAddress" `
            + ", cs-version AS HttpVersion" `
            + ", cs(User-Agent) AS UserAgent" `
            + ", cs(Cookie) AS Cookie" `
            + ", cs(Referer) AS Referrer" `
            + ", cs-host AS Hostname" `
            + ", sc-status AS HttpStatus" `
            + ", sc-substatus AS HttpSubstatus" `
            + ", sc-win32-status AS Win32Status" `
            + ", sc-bytes AS BytesFromServerToClient" `
            + ", cs-bytes AS BytesFromClientToServer" `
            + ", time-taken AS TimeTaken" `
        + " INTO WebsiteLog" `
        + " FROM $httpLogPath\*.log"

    [string] $connectionString = "Driver={SQL Server Native Client 10.0};" `
        + "Server=BEAST;Database=CaelumDW;Trusted_Connection=yes;"

    [string[]] $parameters = @()

    $parameters += $query
    $parameters += "-i:W3C"
    $parameters += "-o:SQL"
    $parameters += "-oConnString:$connectionString"

    Write-Debug "Parameters: $parameters"
    Write-Host "Importing log files to database..."

    & $logParser $parameters
}

function RemoveLogFiles([string] $httpLogPath)
{
    if ([string]::IsNullOrEmpty($httpLogPath) -eq $true)
    {
        throw "The log path must be specified."
    }

    Write-Host "Removing log files..."
    Remove-Item ($httpLogPath + "\*.log")
}

function Main
{
    [string] $httpLogPath = "F:\W3SVC2"

    ExtractLogFiles $httpLogPath
    ImportLogFiles $httpLogPath
    RemoveLogFiles $httpLogPath

    Write-Host -Fore Green "Successfully imported log files."
}

Main
#
# Originally from: "http://sbrickey.com/Tech/Blog/Post/Parsing_IIS_Logs_with_PowerShell"
#

# Define the location of log files and a temporary file
$LogFolder = "C:\inetpub\logs\LogFiles\W3SVC3"
$LogFiles = [System.IO.Directory]::GetFiles($LogFolder, "*.log")
$LogTemp = "C:\inetpub\logs\LogFiles\Temp\AllLogs.tmp"

# Logs will store each line of the log files in an array
$Logs = @()
# Skip the comment lines
$LogFiles | % { Get-Content $_ | where {$_ -notLike "#[D,F,S,V]*" } | % { $Logs += $_ } }
# Then grab the first header line, and adjust its format for later
$LogColumns = ( $LogFiles | select -first 1 | % { Get-Content $_ | where {$_ -Like "#[F]*" } } ) `
              -replace "#Fields: ", "" -replace "-","" -replace "\(","" -replace "\)",""

# Temporarily, store the reformatted logs
Set-Content -LiteralPath $LogTemp -Value ( [System.String]::Format("{0}{1}{2}", $LogColumns, [Environment]::NewLine, ( [System.String]::Join( [Environment]::NewLine, $Logs) ) ) )

# Read the reformatted logs as a CSV file
$Logs = Import-Csv -Path $LogTemp -Delimiter " "

# Sample query : Select all unique users
$Logs | select -Unique csusername
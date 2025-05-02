# Load the Excel file and extract the list of KBs
#$excelFilePath = "C:\LazyWinAdmin\Security Updates 2019-2024.xlsx"
#$worksheetName = '2019'  # Adjust the worksheet name as needed

# Install the ImportExcel module if not already installed
# Install-Module ImportExcel

# Import the module
#Import-Module PSWriteOffice
$Creds = Get-Credential

# Read the list of KBs from the Excel worksheet
#$KBs = Import-OfficeExcel -FilePath $excelFilePath -WorksheetName $worksheetName
#$KBs = $KBs | Select Article -Unique

$KBs = Get-Content C:\LazyWinAdmin\KBs.txt

# Define the list of servers
#$ws = 'WebTeam'
#$Servers = Import-OfficeExcel -FilePath $excelFilePath -WorkSheetName $ws
#$Servers = $Servers | Select EndPoint -Unique

$Servers = @("FBV-ICLGM-P01",
"FBV-MSDEV-D01",
"FBV-PARPA-P01",
"FBV-SCORDEV-D01",
"FBV-SCORDEV-D03",
"FBV-SCORDEV-D04",
"FBV-SCORDEV-D05",
"FBV-SCORDEV-D06",
"FBV-SCORDEV-D08",
"FBV-SCORDEV-D09",
"FBV-SCORDEV-D10",
"FBV-SCORDEV-D12",
"FBV-SCORDEV-D13",
"FBV-SCORDEV-D14",
"FBV-SCORINT-D01",
"FBV-SCRCD10-P03",
"FBV-SCRCD10-P04",
"FBV-SCRCD10-T01",
"FBV-SCRCD10-T02",
"FBV-SCRCM10-T01",
"FBV-SCRXC10-T01",
"FBV-SCSLR10-T01",
"FBV-SCSLR10-T02",
"FBV-SCSLR10-T03",
"FBV-SCSLR10-T04",
"FBV-SONQUBE-P01",
"FBV-SPAPPSS-D01",
"FBV-SPAPPSS-P01",
"FBV-SPAPPSS-T01",
"FBV-SPSRCSS-D01",
"FBV-SPSRCSS-P01",
"FBV-SPSRCSS-T01",
"FBV-SPWFESS-D01",
"FBV-SPWFESS-D02",
"FBV-SPWFESS-P01",
"FBV-SPWFESS-P02",
"FBV-SPWFESS-T01",
"FBV-SPWFESS-T02",
"FBV-UIPATH-D01",
"FBV-UIPATH-P01",
"FBV-UIROBO-P01",
"FBV-WBDV17-D01",
"FBV-WBDV20-D01",
"FBV-WBDV-D02",
"FBV-WBDV-D03",
"FBV-WEBDB-D01",
"FBV-WEBP6-P01")
# Initialize an array to store the results
$results = @()

# Loop through each server
foreach ($server in $Servers) {
    # Initialize a hashtable to store results for the current server
    $serverResult = @{
        ServerName = $server
    }

    # Loop through each KB
    foreach ($KB in $KBs) {
        # Check if the KB is installed on the current server
        $isInstalled = Get-HotFix -ComputerName $server -Credential $Creds | Where-Object {$_.HotFixID -eq $KB}

        # Add the KB and its installation status to the hashtable
        $serverResult[$KB] = if ($isInstalled) { "Installed" } else { "Not Installed" }
    }

    # Add the hashtable to the results array
    $results += New-Object PSObject -Property $serverResult
}

# Display the results
#$results | Format-Table -AutoSize

# Export results to CSV if needed
$results | Export-Csv -Path "C:\LazywinAdmin\KB_Installation_Results.csv" -NoTypeInformation

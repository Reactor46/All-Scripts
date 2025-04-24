# Installing the MS Excel module for PowerShell. This way, PowerShell and Excel can actually interact.
# The if statement ensures that it will only install it if it's not already present.
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}

# Importing the Excel module:
Import-Module ImportExcel

# Establishing where the files that this script will create should go:
$folderPath = "C:\Users\silad\Documents"


# Specify the number of spreadsheets to create
$numSpreadsheets = 5

# For loop to create multiple spreadsheets
for ($i = 1; $i -le $numSpreadsheets; $i++) {
    
    # Define spreadsheet name
    $spreadsheetName = "Spreadsheet_$i.xlsx"

    # Create a new Excel workbook
    $workbook = New-ExcelWorkbook

    # Add a worksheet to the workbook
    $worksheet = $workbook.Worksheets[1]


    # Save the workbook to the specified folder
    $workbook | Export-Excel -Path "$folderPath\$spreadsheetName" -AutoSize -Show
}

Write-Host "$numSpreadsheets new spreadsheet(s) have been created in $folderPath"


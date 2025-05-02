##############################################################################################
# - AD Reporter.ps1
# - Created by Tim Buntrock
# - You need to install RSAT on your admin workstation
##############################################################################################

# Import AD Module 
Import-Module -Name ActiveDirectory 
 
#Create Report folder
New-Item -ItemType directory -Path C:\LazyWinAdmin\AD-Report -force


# Set date variables 
$DateSave = get-date -format d.M.yyyy
$1DayAgo = (Get-Date).AddDays(-1)
$3MonthsAgo = (Get-Date).AddDays(-90)

# Get Domain NetBiosName and save it into a variable
$domobj = get-addomain
$Domain = $domobj.NetBIOSName

# Get Forest Configuration Naming Context and save it into a variable
$RootDSE = [System.DirectoryServices.DirectoryEntry]([ADSI]"LDAP://RootDSE") 
$CfgNC = $RootDSE.Get("configurationNamingContext") 

 
# AD Queries and exports

Write-Host "Finding users that was created in the last 24 hrs" 
Get-ADUser -Filter * -Properties whenCreated | where { $_.whenCreated -ge $1DayAgo } | select SamAccountName,whenCreated |Sort-Object -Property whenCreated | Export-csv "C:\LazyWinAdmin\AD-Report\Users created.csv" -NoType


Write-Host "Finding users with the flag password never expires set" 
Get-ADUser -Filter * -Properties PasswordNeverExpires | where { $_.PasswordNeverExpires -eq $true } | select SamAccountName | Sort-Object -Property SamAccountName | Export-csv "C:\LazyWinAdmin\AD-Report\Users PW Never Expires.csv" -NoType


Write-Host "Finding disabled users" 
Get-ADUser -Filter "Enabled -eq '$false'" | Select SamAccountName | Sort-Object -Property SamAccountName |Sort-Object -Property SamAccountName | Export-csv "C:\LazyWinAdmin\AD-Report\Users disabled.csv" -NoType
    
Write-Host "Finding users that never changed there passwords" 
Get-ADUser –filter * -Properties PasswordLastSet | where { $_.passwordLastSet –eq $null } | Select SamAccountName, enabled | Sort-Object -Property SamAccountName | Export-csv "C:\LazyWinAdmin\AD-Report\Users never changed PW.csv" -NoType


Write-Host "Finding computers that have not logged on for more then 90 days" 
Get-ADComputer -Property Name,lastLogonDate -Filter {lastLogonDate -lt $3MonthsAgo} | Select Name,lastLogonDate | Sort-Object -Property Name | Export-csv "C:\LazyWinAdmin\AD-Report\Computers LastLogon 90d ago.csv" -NoType


Write-Host "Finding disabled computers" 
Get-ADComputer -Property Name -Filter "Enabled -eq '$false'" | Select Name | Sort-Object -Property Name | Export-csv "C:\LazyWinAdmin\AD-Report\Computers Disabled.csv" -NoType


Write-Host "Finding all DCs in your domain"
Get-ADDomainController -Filter * | Select Name | Sort-Object -Property Name | Export-csv "C:\LazyWinAdmin\AD-Report\$Domain Domain Controllers.csv" -NoType


Write-Host "Finding all DHCP servers in your Forest"
Get-ADObject -SearchBase "$CfgNC" -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'" | select name | Sort-Object -Property Name | Export-csv "C:\LazyWinAdmin\AD-Report\Forest DHCP Servers.csv" -NoType


Write-Host "Finding all Subnets with the associated Site and Location name in your Forest"
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$Sites.subnets | Export-Csv "C:\LazyWinAdmin\AD-Report\Forest AD Sites.csv" -NoType


Write-Host "Finding FSMO role holders in your Forest"
Get-ADForest | Select SchemaMaster,DomainNamingMaster | Format-List >"C:\LazyWinAdmin\AD-Report\FSMOs Forest.csv"


Write-Host "Finding FSMO role holders in your Domain"
Get-ADDomain | Select PDCEmulator,RIDMaster,InfrastructureMaster | Format-List >"C:\LazyWinAdmin\AD-Report\FSMOs $Domain Domain.csv"


Write-Host "Finding $Domain PW Policy"
Get-ADDefaultDomainPasswordPolicy >"C:\LazyWinAdmin\AD-Report\$Domain Password Policy.csv"

Write-Host "Finding $Domain GPOs"
Get-GPO -All >"C:\LazyWinAdmin\AD-Report\$Domain GPOs.csv"

Write-Host "Finding $Domain OUs"
Get-ADOrganizationalUnit -Filter * -Properties CanonicalName | Select-Object -Property CanonicalName >"C:\LazyWinAdmin\AD-Report\$Domain OUs.csv"


# Merge all csv files into a excle file and delete them after
$CSVFiles = Get-ChildItem C:\LazyWinAdmin\AD-Report\* -Include *.csv

$CSVFilename = "C:\LazyWinAdmin\AD-Report\AD_Report_$DateSave.xlsx"

Write-Host "Saving AD report to: $CSVFilename"

$excelapp = new-object -comobject Excel.Application
$excelapp.sheetsInNewWorkbook = $CSVFiles.Count
$xlsx = $excelapp.Workbooks.Add()
$Sheet=1

foreach ($CSV in $CSVFiles)
{
$Row=1
$Column=1
$Worksheet = $xlsx.Worksheets.Item($Sheet)
$Worksheet.Name = $CSV.Name
$File = (Get-Content $CSV)
foreach($Line in $File)
{

$Linecontents=$Line

foreach($cell in $Linecontents)
{
$Worksheet.Cells.Item($Row,$Column) = $cell
$Column++
}
$Column=1
$Row++
}
$Sheet++
}

$xlsx.SaveAs($CSVFilename)
$excelapp.quit()

Get-ChildItem C:\LazyWinAdmin\AD-Report\* -Include *.csv | remove-item

# Open report
Invoke-Item $CSVFilename
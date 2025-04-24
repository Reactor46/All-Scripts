$CSV1 = Import-Csv -Path C:\LazyWinAdmin\InstalledKBs.csv
#$CSV1.ComputerName
#$CSV1.HotFixId
#$CSV1.Installed

$servers = $CSV1.ComputerName
$kbInstalled = $CSV1.HotFixId


$CSV2 = Import-Csv -Path C:\LazyWinAdmin\2023_KBs.csv
#$CSV2.CVE
#$CSV2.KB
#$CSV2.Url

$CVE = $CSV2.CVE
$FixedKB = $CSV2.KB
$PatchURL = $CSV2.Url


#"Check ID"	"Title"	"Endpoint"	"Affected Platforms"	"Days Vulnerable"	"Compliance"
$CSV3 = Import-OfficeExcel -FilePath 'C:\LazyWinAdmin\Security Updates 2019-2024.xlsx' -WorkSheetName 'WebTeam'

$CVEVul = $CSV3 | Where {$_.'Check ID' -like "CVE-2023-*"} | Select 'Check ID'
$VulServer = $CSV3.Endpoint


foreach($server in $servers){
    foreach($vul in $VulServer){
        foreach($kb in $kbInstalled){
            

            
       
        
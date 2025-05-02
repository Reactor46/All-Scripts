#Get-DistributionGroup | Select Name, DisplayName | Export-CSV .\All-DISTGROUP.csv -NoTypeInformation
$Groups = Get-Content C:\LazyWinAdmin\EXCHANGE\DGList.txt | Get-DistributionGroup
$Groups | ForEach-Object {
$members = ''
Get-DistributionGroupMember $group | ForEach-Object {
        If($members) {
              $members=$members + ";" + $_.Name
           } Else {
              $members=$_.Name
           }
  }
New-Object -TypeName PSObject -Property @{
      GroupName = $group
      Members = $members
     }
} | Export-CSV "C:\LazyWinAdmin\EXCHANGE\ALL-DIST-GROUP-MEMBERS.csv" -NoTypeInformation -Encoding UTF8
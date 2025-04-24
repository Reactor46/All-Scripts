Set-ExecutionPolicy Unrestricted
  
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue 

Get-SPWebApplication | Get-SPSite -Limit "All" | Where-Object {$_.URL -NotLike "*MySites*"} | Get-SPWeb -Limit "All" | Sort-Object url | Export-CSV -Path "D:\UpgradeInfo\AllFarmWebs.csv" 
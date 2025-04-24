
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop 

Get-SPSite https://pulse.kscpulse.com/hcf/home | Get-SPSite -Limit ALL |%{ Get-SPFeature -Site $_ } | Select DisplayName,ID -Unique
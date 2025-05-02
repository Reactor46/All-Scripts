
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

Get-spwebapplication -includecentraladministration |
where {$_.DisplayName -match "SharePoint Central Administration*"} |
 select DisplayName,Url
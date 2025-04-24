#Import-Module Webadministration
#Get-ChildItem -Path IIS:\Sites | Out-File "\\fbv-wbdv20-d01\D$\IIS-Servers\$env:COMPUTERNAME.Sites.csv"

$sites=[xml](c:\windows\system32\inetsrv\appcmd.exe list site  /xml)
$sites.appcmd.site | Select SITE.NAME, bindings | Export-Csv -NoTypeInformation "\\fbv-wbdv20-d01\D$\IIS-Servers\$env:COMPUTERNAME.Sites.csv"
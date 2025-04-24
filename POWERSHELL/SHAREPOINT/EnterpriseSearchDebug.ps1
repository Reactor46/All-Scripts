Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue


$ssa = Get-SPEnterpriseSearchServiceApplication -Identity "Search Service Application" 
Suspend-SPEnterpriseSearchServiceApplication -Identity $ssa
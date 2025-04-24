Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

Get-SPDatabase | where {$_.Type -notcontains "Content Database" -and `
$_.Type -notcontains "Configuration Database"} | sort Type | format-table –autosize
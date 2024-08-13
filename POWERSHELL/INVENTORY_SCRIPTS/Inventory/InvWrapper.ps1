$ScriptRoot = "C:\\Inventory\ServerInventoryReport"

Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain Contoso"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain PHX"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain BIZ"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain TST"
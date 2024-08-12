$ScriptRoot = "C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\Inventory\ServerInventoryReport"

Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain FNBM"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain PHX"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain BIZ"
Invoke-Expression "$ScriptRoot\GatherData.ps1 -Domain TST"
Function Show-MemoryUsage {
 
[cmdletbinding()]
Param()
 
#get memory usage data
$data = Test-MemoryUsage
 
Switch ($data.Status) {
"OK" { $color = "Green" }
"Warning" { $color = "Yellow" }
"Critical" {$color = "Red" }
}
 
$title = @"
 
Memory Check
------------
"@
 
Write-Host $title -foregroundColor Cyan
 
$data | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor $color
 
}
 
#set-alias -Name smu -Value Show-MemoryUsage
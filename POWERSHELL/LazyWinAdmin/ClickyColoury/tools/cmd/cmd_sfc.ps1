# api: multitool
# version: 1.0
# title: SFC /SCANNOW
# description: scan protected system files (and replace)
# type: window
# category: cmd
# hidden: 0
# key: 5|sfc
# config: -
#
# runs sfc /scannow on remote computer

Param($machine = (Read-Host "Computer"))

Write-Host "Starting SFC /SCANNOW..."
psexec \\$machine sfc /scannow

Read-Host "---END---"

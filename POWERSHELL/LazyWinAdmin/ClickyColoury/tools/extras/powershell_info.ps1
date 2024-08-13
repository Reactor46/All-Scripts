# api: multitool
# version: 0.0
# title: PowerShell info
# description: runtime details
# type: inline
# category: misc
# hidden: 1
# icon: powershell
# config: -
#
# Some PSHost details

Write-Host -f Yellow "❏ Host"
($Host |FL | out-string).trim() |write-host

Write-Host -f Yellow "❏ VerTbl"
($PSVersionTable |Ft | out-string).trim() |write-host

Write-Host -f Yellow "❏ [Accels]"
([psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")::get | fT -auto -wrap | out-string -width 60).trim() | write-host

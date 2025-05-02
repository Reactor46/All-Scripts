'By K8L0

Const SHOW_ACTIVE_APP = 1
Set objShell = Wscript.CreateObject("Wscript.Shell")
objShell.Run ("LogonSecBanner.hta"), SHOW_ACTIVE_APP, True
Wscript.Quit
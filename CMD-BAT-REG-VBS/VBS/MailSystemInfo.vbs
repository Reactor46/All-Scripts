Const olByValue = 1
Const olMailItem = 0

Dim oOApp 
Dim oOMail

Set Sh = CreateObject("WScript.Shell") 
Set Exec = Sh.Exec("%comspec% /c ipconfig /all") 
If Not Exec.StdOut.AtEndOfStream Then T = Exec.StdOut.ReadAll 
' MsgBox T

Set WshNetwork = WScript.CreateObject("WScript.Network")
Set oOApp = CreateObject("Outlook.Application")
Set oOMail = oOApp.CreateItem(olMailItem)

Set oOApp = CreateObject("Outlook.Application")
Set oOMail = oOApp.CreateItem(olMailItem)

oOMail.display ' Fenster anzeigen
oOMail.To = "<email@yourDomain.com>"
' SAMPLE: oOMail.To = "a@b.com"
oOMail.Subject = "Systemkonfiguration von " & WshNetwork.UserName & " auf " & WshNetwork.ComputerName
oOMail.Body = "Computer Name = "  & WshNetwork.ComputerName & VBCrLf _
				& "User Name = " & WshNetwork.UserName & VBCrLf _
				& T
'oOMail.Attachments.Add "c:\"& WshNetwork.ComputerName & "_"  & WshNetwork.UserName & ".txt", olByValue, 1

'oOMail.Send '<< dieser Befehl würde die Warnung anzeigen

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate oOMail
WshShell.SendKeys("%s") ' Sende ein "Alt-S". Outlook denkt der User "sendet

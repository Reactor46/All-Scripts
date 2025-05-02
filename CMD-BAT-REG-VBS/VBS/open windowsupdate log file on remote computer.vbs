'This script will get a remote computers windowsupdate.log file
'it also changes the association for .log files to open with notepad

'Get and open WindowsUpdate log file on remote Computer
on error resume next

Set WshShell = Wscript.CreateObject("Wscript.Shell") 
strcomputer	  = inputbox("Enter remote computer name or leave as localhost for this computer","Get WindowsUpdate.log file","Localhost")

If strComputer = "" Then
  WScript.Quit
End If

'Associate .log Files with Notepad
wshShell.run "%comspec% /c c: & assoc .log=txtfile"
WScript.Sleep 500 'give association time to take effect

'open WindowsUpdate.log file on remote computer
wshShell.Run "\\" & strcomputer & "\c$\Windows\WindowsUpdate.log"











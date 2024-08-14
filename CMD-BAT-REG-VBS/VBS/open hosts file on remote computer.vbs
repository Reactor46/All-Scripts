'This script will get and allow you to edit a remote computers hosts file
'it also changes the association for extensionless files like the hosts file to open with notepad

'Get and open Remote Computers Hosts file
on error resume next

Set WshShell = Wscript.CreateObject("Wscript.Shell") 
strcomputer	  = inputbox("Enter remote computer name or leave as localhost for this computer","Get Hosts file","Localhost")

If strComputer = "" Then
  WScript.Quit
End If

'Associate Extensionless Files with Notepad
wshShell.run "%comspec% /c c: & assoc .=txtfile"
WScript.Sleep 500 'give association time to take effect

'open hosts file on remote
wshShell.Run "\\" & strcomputer & "\c$\Windows\System32\drivers\etc\hosts"











On Error Resume Next

Set objShell = CreateObject( "Wscript.Shell" )

' The "True" argument will make the script wait for the screensaver to exit

returnVal = objShell.Run( "Disable_USB_Storage.reg" , 1, False)



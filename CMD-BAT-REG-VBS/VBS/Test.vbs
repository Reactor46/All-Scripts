If DEBUGMODE Then 
	WScript.Echo "ran test.vbs"
		WScript.Echo config.key("installstring")
Else
	Call Terminate(0)
End If
If Err.Number <> 0 Then
	WScript.Echo Err.Number & " " & Err.Description & " " & Err.Source
	WScript.Sleep 15000
End If

Call logging.write(config.key("installstring"),1)
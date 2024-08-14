result=Msgbox("Click Yes to Return Exit Code of '0'" & _
vbCrLf & "Click No to Return Exit Code of '1'" & _
vbCrLf & "Click Cancel to Return Exit Code of '1238'", _
vbYesNoCancel+vbQuestion, "Dummy Program")

if result = 6 then
	wscript.quit 0
elseif result = 7 then
	wscript.quit 1
elseif result = 2 then
	wscript.quit 1238
end if



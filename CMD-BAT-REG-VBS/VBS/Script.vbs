Dim objShell
Dim IE
Dim str_date
Dim Cmd
Dim d
Dim m
Dim y
Set objShell = Wscript.CreateObject("Wscript.Shell")
Set WSHNetwork = WScript.CreateObject("WScript.Network")
Dim strstatus

On error resume next

Call CreateIE()
'showstat("Start on: " & Time())

d = Day(date())
m = Month(date())
y = Year(date())

if (m = 1) then
	m = 12
	y = y - 1
else
	m = m - 1
	if (d = 31) then
		m = m - 1
	end if
end if

if (m = 2) then
	if (d = 29) then
		Cmd = "7z.exe a c:\" &y & "01" &d & ".7z c:\" & y & "01" & d & "*.*"
		showstat(Cmd)
		objShell.Run Cmd
	end if
	if (d = 30) then
		m = m - 1
	end if
end if

if (m<10) then
	m = "0"  &m
end if

if (d<10) then
	d = "0"  &d
end if

Cmd = "7z.exe a c:\" &y &m &d & ".7z c:\" & y & m & d & "*.*"

showstat(Cmd)
objShell.Run Cmd

Sub CreateIE()
	On Error Resume Next
        Set IE = CreateObject("InternetExplorer.Application")
        With IE
        	.navigate "C:\Script.htm"
                .resizable=0
                .height=430
                .width=400
                .menubar=0
                .toolbar=0
                .statusBar=0
                .visible=1
  	End With
        Do while IE.Busy
        	Wscript.Sleep 100
        Loop
End Sub


Sub Main()
    On Error Resume Next
End Sub


Function showstat(strmessage)
	strstatus=strmessage + VBCRLF + strstatus 
	ie.document.all.wstatus.InnerText = strstatus
end function
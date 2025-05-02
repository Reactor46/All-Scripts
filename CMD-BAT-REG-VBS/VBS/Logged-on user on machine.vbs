
strComputer = InputBox ("Enter Machine Name")


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo "Logged-on user: " & objComputer.UserName
Next

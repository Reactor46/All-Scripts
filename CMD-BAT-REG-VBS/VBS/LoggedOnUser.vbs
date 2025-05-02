strAnswer = InputBox("Please enter computer Name:", "Logged on user")

strComputer = strAnswer
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set colComputer = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo "Logged-on user: " & objComputer.UserName
Next


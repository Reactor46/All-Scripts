strComputer = "."
Set objUser = GetObject("WinNT://" & strComputer & "/Administrator, user")
objUser.SetPassword "WTF2011g0Fy$"
objUser.SetInfo
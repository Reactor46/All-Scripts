Dim strName

Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

' objUser.name returns CN=LAST.FIRST.MI.#.N
strName = Split(objUser.name, "=")(1)
WScript.Echo strName
' =========================================================================================================
' Determine Which Domain Controller Authenticated The Domain User and The Domain-Joined Computer
' Also Know The Site Name of The Computer
' =========================================================================================================

Option Explicit

Dim ObjRootDSE, StrDC, ObjADSysInfo, StrNet

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDC = ObjRootDSE.Get("dnsHostName")
WScript.Echo "Authenticating Domain Controller: " & Trim(StrDC)
Set ObjRootDSE = Nothing

' --- List the Site Name of the Local Computer

Set ObjADSysInfo = CreateObject("ADSystemInfo")
WScript.Echo "Local Computer's Current Site name: " & Trim(ObjADSysInfo.SiteName)
Set ObjADSysInfo = Nothing
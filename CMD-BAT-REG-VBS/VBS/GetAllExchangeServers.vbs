' ==========================================================================================
' VBScript: Get a List Of All Exchange Servers; All Version of Exchange Servers
' This Script will work even if there is a mixed Exchange environment
' For Example, a Mix of Exchange 2003, Exchange 2007, Exchange 2010 and Exchange 2013
' ==========================================================================================

Option Explicit

Dim ObjConn, ObjRS, ObjRootDSE, ObjCommand
Dim StrConfigNC, StrDomName

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrConfigNC = Trim(ObjRootDSE.Get("ConfigurationNamingContext"))
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjCommand = CreateObject("ADODB.Command"):	ObjCommand.ActiveConnection = ObjConn
ObjCommand.CommandText = "<LDAP://" & StrConfigNC & ">;(ObjectCategory=msExchExchangeServer);Name;Subtree"
Set ObjRS = ObjCommand.Execute
If Not ObjRS.EOF Then
	ObjRS.MoveLast:	ObjRS.MoveFirst
	WScript.Echo "Total Exchange Servers: " & ObjRS.RecordCount & VbCrLf
	While Not ObjRS.EOF
		WScript.Echo "Exchange Server: " & Trim(ObjRS.Fields(0).Value)
		' -- If you want to know the Site of Each Exchange Server, uncomment the line below AND Also uncomment the
		' -- Sub, named NowGetExchangeServerSite, which has been commented below		
		' -- NowGetExchangeServerSite(Trim(ObjRS.Fields(0).Value))
		ObjRS.MoveNext
	Wend
End If
WScript.Echo vbNullString
Set ObjRS = Nothing:	Set ObjCommand = Nothing
ObjConn.Close:	Set ObjConn = Nothing
WScript.Quit

' -- Private Sub NowGetExchangeServerSite (ExchName)
	' -- On Error Resume Next
	' -- Dim ObjADSysInfo
	' -- Set ObjADSysInfo = CreateObject("ADSystemInfo")
    ' -- WScript.Echo "Site Name: " & ObjADSysInfo.GetDCSiteName(ExchName)
    ' -- Set ObjADSysInfo = Nothing
' -- End Sub
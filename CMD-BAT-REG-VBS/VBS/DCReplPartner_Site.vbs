' =========================================================================================================
' Get All Domain Controllers in One Site and Their Replication Partners

' If executed from a Remote Machine, User must have priviledge to connect to DC remotely via RPC
' Typically -- Domain Admin Privilege

' Usage: CScript DCReplPartner_Site.vbs
' =========================================================================================================

Option Explicit

Dim ObjADSysInfo, ObjConn, ObjRS, ObjRootDSE
Dim ObjNTDS, ObjDC, StrForestRootDN, StrDomainDN
Dim StrForestRoot, StrADSite, StrADsPath, StrFilter


' -- Enter The AD Site-Name Here
StrADSite = "Mumbai"

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrForestRootDN = Trim(ObjRootDSE.Get("RootDomainNamingContext"))
StrDomainDN = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing
Set ObjADSysInfo = CreateObject("ADSystemInfo")
WScript.Echo vbNullString
WScript.Echo "Forest Root: " & Trim(ObjADSysInfo.ForestDNSName) & VbCrLf
StrForestRoot = Trim(ObjADSysInfo.ForestDNSName)
Set ObjADSysInfo = Nothing

Set ObjRootDSE = GetObject("LDAP://" & StrForestRoot & "/RootDSE")
StrADsPath = "<LDAP://CN=Servers,CN=" & StrADSite & ",CN=Sites," & Trim(ObjRootDSE.Get("ConfigurationNamingContext")) & ">;"
Set ObjRootDSE = Nothing
StrFilter  = "(Objectcategory=NTDSDSA);"
Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrADsPath & StrFilter & "DistinguishedName;" & "Subtree", ObjConn
If Not ObjRS.EOF Then
	ObjRS.MoveFirst
	WScript.Echo "Domain Controllers in Site " & StrADSite & " are:"
	While Not ObjRS.EOF
		Set ObjNTDS = GetObject("LDAP://" & Trim(ObjRS.Fields(0).Value))
		Set ObjDC = GetObject( ObjNTDS.Parent )
		WScript.Echo "**** " & Trim(ObjDC.Get("DNSHostName")) & " ****"
		WScript.Echo "----------------------------------------"
		WScript.Echo vbTab & "This DC's Replication Partners are: "
		WScript.Echo vbTab & "======================================"
		' --- Now Get The Peplication Partnet of Each DC in Site
		GetReplicationPartnerOfDC(Trim(ObjDC.Get("DNSHostName")))
		Set ObjDC = Nothing:	Set ObjNTDS = Nothing		
		WScript.Echo vbNullString:	ObjRS.MoveNext
	Wend
Else
	WScript.Echo "There is No Domain Controller in Site: " & StrADSite
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

Private Sub GetReplicationPartnerOfDC(ThisDC)

	Dim ObjWMI, ColReplPartitions, ObjReplJob
	Dim CurPos, NewString
	
	Set ObjWMI = GetObject("WinMgmts:" & "{ImpersonationLevel=Impersonate}!\\" & ThisDC & "\Root\MicrosoftActiveDirectory")
	Set ColReplPartitions = ObjWMI.ExecQuery ("Select * from MSAD_ReplNeighbor")
	For Each ObjReplJob in ColReplPartitions 
		' -- Domain Partition
		If StrComp(Trim(ObjReplJob.NamingContextDN), StrDomainDN, vbTextCompare) = 0 Then
			If InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers") > 0 Then
				CurPos = InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers")
				NewString = Trim(Left(Trim(ObjReplJob.SourceDSADN), Curpos-2))
				NewString = Right(NewString, Len(NewString)-20)
				GetDCFQDN "Domain", NewString
				WScript.Echo vbTab & "DN LDAP Path: " & Trim(ObjReplJob.SourceDSADN)
				WScript.Echo vbNullString
			End If
		End If
		' -- Configuration Partition
		If InStr(Trim(ObjReplJob.NamingContextDN), "CN=Configuration,DC=") > 0 AND InStr(Trim(ObjReplJob.NamingContextDN), "CN=Schema") = 0 Then
			If InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers") > 0 Then
				CurPos = InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers")
				NewString = Trim(Left(Trim(ObjReplJob.SourceDSADN), Curpos-2))
				NewString = Right(NewString, Len(NewString)-20)
				GetDCFQDN "Configuration", NewString
				WScript.Echo vbTab & "DN LDAP Path: " & Trim(ObjReplJob.SourceDSADN)
				WScript.Echo vbNullString
			End If
		End If
		' -- Schema Partition
		If InStr(Trim(ObjReplJob.NamingContextDN), "CN=Schema") > 0 Then
			If InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers") > 0 Then
				CurPos = InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers")
				NewString = Trim(Left(Trim(ObjReplJob.SourceDSADN), Curpos-2))
				NewString = Right(NewString, Len(NewString)-20)
				GetDCFQDN "Schema", NewString
				WScript.Echo vbTab & "DN LDAP Path: " & Trim(ObjReplJob.SourceDSADN)
				WScript.Echo vbNullString
			End If
		End If
		' -- Forest DNS Zones
		If InStr(Trim(ObjReplJob.NamingContextDN), "DC=ForestDnsZones") > 0 Then
			If InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers") > 0 Then
				CurPos = InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers")
				NewString = Trim(Left(Trim(ObjReplJob.SourceDSADN), Curpos-2))
				NewString = Right(NewString, Len(NewString)-20)
				GetDCFQDN "ForestDNSZones", NewString
				WScript.Echo vbTab & "DN LDAP Path: " & Trim(ObjReplJob.SourceDSADN)
				WScript.Echo vbNullString
			End If
		End If		
		' -- Domain DNS Zones
		If InStr(Trim(ObjReplJob.NamingContextDN), "DC=DomainDnsZones") > 0 Then
			If InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers") > 0 Then
				CurPos = InStr(Trim(ObjReplJob.SourceDSADN), "CN=Servers")
				NewString = Trim(Left(Trim(ObjReplJob.SourceDSADN), Curpos-2))
				NewString = Right(NewString, Len(NewString)-20)
				GetDCFQDN "DomainDNSZones", NewString
				WScript.Echo vbTab & "DN LDAP Path: " & Trim(ObjReplJob.SourceDSADN)
				WScript.Echo vbNullString
			End If
		End If		
	Next
End Sub

Private Sub GetDCFQDN(Partition, ThisDC)

	Dim StrDomName, StrSQL
	Dim ObjNewConn, ObjNewRS, ObjDC
	
	Set ObjRootDSE = GetObject("LDAP://RootDSE")
	StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
	Set ObjRootDSE = Nothing
	StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'Computer' AND Name = '" & ThisDC & "'"
	Set ObjNewConn = CreateObject("ADODB.Connection")
	ObjNewConn.Provider = "ADsDSOObject":	ObjNewConn.Open "Active Directory Provider"
	Set ObjNewRS = CreateObject("ADODB.Recordset")
	ObjNewRS.Open StrSQL, ObjNewConn
	If Not ObjNewRS.EOF Then
		Set ObjDC = GetObject(Trim(ObjNewRS.Fields("ADsPath").Value))
		If StrComp(Partition, "Schema", vbTextCompare) = 0 Then
			WScript.Echo vbTab & "Schema Partition: " & Trim(ObjDC.DNSHostName)
		End If
		If StrComp(Partition, "Configuration", vbTextCompare) = 0 Then
			WScript.Echo vbTab & "Configuration Partition: " & Trim(ObjDC.DNSHostName)
		End If
		If StrComp(Partition, "Domain", vbTextCompare) = 0 Then
			WScript.Echo vbTab & "Domain Partition: " & Trim(ObjDC.DNSHostName)
		End If
		If StrComp(Partition, "ForestDNSZones", vbTextCompare) = 0 Then
			WScript.Echo vbTab & "Forest DNS Zones Partition: " & Trim(ObjDC.DNSHostName)
		End If
		If StrComp(Partition, "DomainDNSZones", vbTextCompare) = 0 Then
			WScript.Echo vbTab & "Domain DNS Zones Partition: " & Trim(ObjDC.DNSHostName)
		End If
		Set ObjDC = Nothing
	End If
	ObjNewRS.Close:	Set ObjNewRS = Nothing
	ObjNewConn.Close:	Set ObjNewConn = Nothing
End Sub
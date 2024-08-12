'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 04/11/2006
' DomainInfo.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' Output Format 
Const HTML = 1 ' Save as HTML (if Equals 1)
Const TXT = 0 ' Save as Text (if Equals 1)
' Log File Name and PAth
Const FilePath = "C:\"
Const FileName = "Domain Log"
Const ADS_SCOPE_SUBTREE = 2
' Name of one of your DC you want to scan
Const DC_NAME = "DC01"

' Set the Log File Path
If Txt = 1 Then
	FileLoc = FilePath & FileName & ".txt"
End If
If HTML = 1 Then
	FileLoc = FilePath & FileName & ".html"
End If
If Txt = 1 AND HTML = 1 Then
	FileLoc = FilePath & FileName & ".txt"
End If

'  Get The Configuration Naming Context
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
ArrDCName = Split(strConfigurationNC,",")
strDCPath = ArrDCName(1) & "," & ArrDCName(2)

Sub WriteText (strText, Header, HTMLTag)
	'On Error Resume Next
	If objFSO.FileExists(FileLoc) Then
	    Set objFile = objFSO.OpenTextFile(FileLoc,ForAppending)
	Else
	    Set objFile = objFSO.CreateTextFile(FileLoc)
	    If HTML = 1 Then
			objFile.WriteLine "<HTML><Head><Title>" & FileName & "</Title></Head>"
			objFile.WriteLine "<Body>"
		End If
	End If
	
	If HTML = 1 Then
		If Header <> "" Then
			objFile.Write "<" & Header & ">"
		End If
		If HTMLTag <> "" Then
			If HTMLTag = "BR" Then
				'Do Nothing
			Else
				objFile.Write "<" & HTMLTag & ">"
			End If
		End If
	End If
	objFile.WriteLine strText
	If HTML = 1 Then
		If Header <> "" Then
			objFile.Write "</" & Header & ">"
		End If
		If HTMLTag <> "" Then
			If HTMLTag = "BR" Then
				objFile.Write "<" & HTMLTag & ">"
			Else
				objFile.WriteLine "</" & HTMLTag & ">"
			End If
		End If
	End If
	'objFile.WriteLine
End Sub

Sub CloseText()
	Set objFile = objFSO.OpenTextFile(FileLoc,ForAppending)
	If HTML = 1 Then
		objFile.Write "</Body></HTML>"
	End If
	objFile.Close
End Sub
	
Sub DomainControllers()
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	
	objCommand.CommandText = _
	    "SELECT ADsPath FROM 'LDAP://" & strConfigurationNC & "' WHERE objectClass='nTDSDSA'"  
	Set objRecordSet = objCommand.Execute
	
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
	    Set objParent = GetObject(GetObject(objRecordset.Fields("ADsPath")).Parent)
	    WriteText objParent.CN,"","BR"
	    objRecordSet.MoveNext
	Loop
End Sub

Sub PasswordProperty(strDCName)
	Set objHash = CreateObject("Scripting.Dictionary")
	 
	objHash.Add "DOMAIN_PASSWORD_COMPLEX", &h1
	objHash.Add "DOMAIN_PASSWORD_NO_ANON_CHANGE", &h2
	objHash.Add "DOMAIN_PASSWORD_NO_CLEAR_CHANGE", &h4
	objHash.Add "DOMAIN_LOCKOUT_ADMINS", &h8
	objHash.Add "DOMAIN_PASSWORD_STORE_CLEARTEXT", &h16
	objHash.Add "DOMAIN_REFUSE_PASSWORD_CHANGE", &h32
	 
	Set objDomain = GetObject("LDAP://" & strDCName)
	 
	intPwdProperties = objDomain.Get("PwdProperties")
	WriteText "Password Properties = ","","B"
	WriteText intPwdProperties,"","BR"
	 
	For Each Key In objHash.Keys
	    If objHash(Key) And intPwdProperties Then 
	        WriteText Key & " is enabled","","BR"
	    Else
	        WriteText Key & " is disabled","","BR"
	    End If
	Next
End Sub

Sub PasswordPolicy(strDCName)
	Const MIN_IN_DAY = 1440
	Const SEC_IN_MIN = 60
	 
	Set objDomain = GetObject("WinNT://" & strDCName)
	Set objAdS = GetObject("LDAP://" & strDCName)
	 
	intMaxPwdAgeSeconds = objDomain.Get("MaxPasswordAge")
	intMinPwdAgeSeconds = objDomain.Get("MinPasswordAge")
	intLockOutObservationWindowSeconds = objDomain.Get("LockoutObservationInterval")
	intLockoutDurationSeconds = objDomain.Get("AutoUnlockInterval")
	intMinPwdLength = objAds.Get("minPwdLength")
	 
	intPwdHistoryLength = objAds.Get("pwdHistoryLength")
	intPwdProperties = objAds.Get("pwdProperties")
	intLockoutThreshold = objAds.Get("lockoutThreshold")
	intMaxPwdAgeDays = _
	  ((intMaxPwdAgeSeconds/SEC_IN_MIN)/MIN_IN_DAY) & " days"
	intMinPwdAgeDays = _
	  ((intMinPwdAgeSeconds/SEC_IN_MIN)/MIN_IN_DAY) & " days"
	intLockOutObservationWindowMinutes = _
	  (intLockOutObservationWindowSeconds/SEC_IN_MIN) & " minutes"
	 
	If intLockoutDurationSeconds <> -1 Then
	  intLockoutDurationMinutes = _
	(intLockOutDurationSeconds/SEC_IN_MIN) & " minutes"
	Else
	  intLockoutDurationMinutes = _
	    "Administrator must manually unlock locked accounts"
	End If
	 
	WriteText "maxPwdAge = " & intMaxPwdAgeDays,"","BR"
	WriteText "minPwdAge = " & intMinPwdAgeDays,"","BR"
	WriteText "minPwdLength = " & intMinPwdLength,"","BR"
	WriteText "pwdHistoryLength = " & intPwdHistoryLength,"","BR"
	WriteText "pwdProperties = " & intPwdProperties,"","BR"
	WriteText "lockOutThreshold = " & intLockoutThreshold,"","BR"
	WriteText "lockOutObservationWindow = " & intLockOutObservationWindowMinutes,"","BR"
	WriteText "lockOutDuration = " & intLockoutDurationMinutes,"","BR"
End Sub

Sub SitesNServers()
	On Error Resume Next
	 
	strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
	Set objSitesContainer = GetObject(strSitesContainer)
	objSitesContainer.Filter = Array("site")
	 
	For Each objSite In objSitesContainer
	    WriteText objSite.CN,"","B"
	    strSiteName = objSite.Name
	    strServerPath = "LDAP://cn=Servers," & strSiteName & ",cn=Sites," & strConfigurationNC
	    Set colServers = GetObject(strServerPath)
	 
	    For Each objServer In colServers
	        WriteText vbTab & objServer.CN,"","BR"
	    Next
	    WriteText "","","BR"
	Next
End Sub

Sub TrustDomain(strComputer)
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & _
	        strComputer & "\root\MicrosoftActiveDirectory")
	
	Set colTrustList = objWMIService.ExecQuery _
	    ("Select * from Microsoft_DomainTrustStatus")
	
	For each objTrust in colTrustList
	    WriteText "Trusted domain: ","","B"
	    WriteText objTrust.TrustedDomain,"","BR"
	    WriteText "Trust direction: ","","B"
	    WriteText objTrust.TrustDirection,"","BR"
	    WriteText "Trust type: ","","B" 
	    WriteText objTrust.TrustType,"","BR"
	    WriteText "Trust attributes: ","","B"
	    WriteText objTrust.TrustAttributes,"","BR"
	    WriteText "Trusted domain controller name: ","","B"
	    WriteText objTrust.TrustedDCName,"","BR"
	    WriteText "Trust status: ","","B" 
	    WriteText objTrust.TrustStatus,"","BR"
	    WriteText "Trust is OK: ","","B"
	    WriteText objTrust.TrustIsOK,"","BR"
	Next
End Sub

Sub ReplicationPartners(strComputer)
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & _
	        strComputer & "\root\MicrosoftActiveDirectory")
	Set colReplicationOperations = objWMIService.ExecQuery _
	    ("Select * from MSAD_ReplNeighbor")
	For each objReplicationJob in colReplicationOperations 
	    WriteText objReplicationJob.Domain,"","BR"
	    WriteText objReplicationJob.NamingContextDN,"","BR"
	    WriteText objReplicationJob.SourceDsaDN,"","BR"
	    WriteText objReplicationJob.LastSyncResult,"","BR"
	    WriteText objReplicationJob.NumConsecutiveSyncFailures,"","BR"
	Next
End Sub

Sub DomainInfo(strComputer)
	On Error Resume Next
	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	Set colItems = objWMIService.ExecQuery("Select * from Win32_NTDomain")
	
	For Each objItem in colItems
	    WriteText "Client Site Name: ","","B"
	    WriteText objItem.ClientSiteName,"","BR"
	    WriteText "DC Site Name: ","","B"
	    WriteText objItem.DcSiteName,"","BR"
	    WriteText "Description: ","","B"
	    WriteText objItem.Description,"","BR"
	    WriteText "DNS Forest Name: ","","B"
	    WriteText objItem.DnsForestName,"","BR"
	    WriteText "Domain Controller Address: ","","B"
	    WriteText objItem.DomainControllerAddress,"","BR"
	    WriteText "Domain Controller Address Type: ","","B"
	    WriteText objItem.DomainControllerAddressType,"",""
	    If objItem.DomainControllerAddressType = 1 Then
			WriteText "(IP)","","BR"
		End If
	    WriteText "Domain Controller Name: ","","B"
	    WriteText objItem.DomainControllerName,"","BR"
	    WriteText "Domain GUID: ","","B"
	    WriteText objItem.DomainGuid,"","BR"
	    WriteText "Domain Name: ","","B"
	    WriteText objItem.DomainName,"","BR"
	    WriteText "DS Directory Service Flag: ","","B"
	    WriteText objItem.DSDirectoryServiceFlag,"","BR"
	    WriteText "DS DNS Controller Flag: ","","B"
	    WriteText objItem.DSDnsControllerFlag,"","BR"
	    WriteText "DS DNS Domain Flag: ","","B"
	    WriteText objItem.DSDnsDomainFlag,"","BR"
	    WriteText "DS DNS Forest Flag: ","","B"
	    WriteText objItem.DSDnsForestFlag,"","BR"
	    WriteText "DS Global Catalog Flag: ","","B"
	    WriteText objItem.DSGlobalCatalogFlag,"","BR"
	    WriteText "DS Kerberos Distribution Center Flag: ","","B"
	    WriteText objItem.DSKerberosDistributionCenterFlag,"","BR"
	    WriteText "DS Primary Domain Controller Flag: ","","B"
	    WriteText objItem.DSPrimaryDomainControllerFlag,"","BR"
	    WriteText "DS Time Service Flag: ","","B"
	    WriteText objItem.DSTimeServiceFlag,"","BR"
	    WriteText "DS Writable Flag: ","","B"
	    WriteText objItem.DSWritableFlag,"","BR"
	    WriteText "Name: ","","B"
	    WriteText objItem.Name,"","BR"
	    WriteText "Primary Owner Contact: ","","B"
	    WriteText objItem.PrimaryOwnerContact,"","BR"
	    
	Next
End Sub

Sub AllDisabledUsers(strDCName)
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	
	objCommand.CommandText = _
	    "<LDAP://" & strDCName & ">;(&(objectCategory=User)" & _
	        "(userAccountControl:1.2.840.113556.1.4.803:=2));Name;Subtree"  
	Set objRecordSet = objCommand.Execute
	
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
	    WriteText objRecordSet.Fields("Name").Value,"","BR"
	    objRecordSet.MoveNext
	Loop
End Sub


Sub CountObjects(ADObject)
	'On Error Resume Next
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	
	objCommand.CommandText = _
	    "SELECT Name FROM 'LDAP://" & strDCName & "' WHERE objectCategory='" & ADObject & "'"  
	Set objRecordSet = objCommand.Execute
	
	WriteText objRecordSet.RecordCount,"","BR"
End Sub

' Code Starts Here
WriteText "Domain Info","H1","B"
DomainInfo DC_NAME
WriteText "Domain Controllers","H1","B"
DomainControllers
WriteText "Password Policy and Properties","H1","B"
PasswordPolicy strDCPath
WriteText "","","BR"
PasswordProperty DC_NAME
WriteText "Domain Sites","H1","B"
SitesNServers
WriteText "Domain Trusts","H1","B"
TrustDomain DC_NAME
WriteText "Replication Partners","H1","B"
ReplicationPartners DC_NAME
WriteText "All Disabled Users","H1","B"
AllDisabledUsers strDCPath
WriteText "All Users","H1","B"
CountObjects "user"
WriteText "All Computers","H1","B"
CountObjects "computer" 
WriteText "All organizational Units","H1","B"
CountObjects "organizationalUnit"
CloseText

WScript.Echo "Done!"
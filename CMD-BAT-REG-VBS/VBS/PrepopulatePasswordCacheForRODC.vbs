'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

Option Explicit

Const WshRunning = 0
Dim objShell,objExec,strOutput,objRegExp,objADInfo,objRootDSE,strDefaultNamingContext
Dim arrGroupMemberDNs(),i

Set objADInfo = CreateObject("ADSystemInfo")
Set objRootDSE = GetObject("LDAP://RootDSE")
Set objShell = CreateObject("WScript.Shell")
Set objRegExp = New RegExp
objRegExp.Global = True
objRegExp.IgnoreCase = True
strDefaultNamingContext = objRootDSE.Get("defaultNamingContext")

Function TestOSCUserPrivilege
	'This function is used to check the privilege of current user.
	Dim strHighMandatoryLevel,blnIsElevated,blnIsDomainAdmin,strDomainAdmins
	strHighMandatoryLevel = "S-1-16-12288"
	strDomainAdmins = "Domain Admins"
	Set objExec = objShell.Exec("whoami /all")
	Do While (objExec.Status <> WshRunning)
		Call WScript.Sleep(200)
	Loop
	strOutput = objExec.StdOut.ReadAll()
	objRegExp.Pattern = strHighMandatoryLevel
	blnIsElevated = objRegExp.Test(strOutput)
	objRegExp.Pattern = strDomainAdmins
	blnIsDomainAdmin = objRegExp.Test(strOutput)	
	TestOSCUserPrivilege = blnIsElevated And blnIsDomainAdmin
End Function

Function TestOSCIsRODC(DCName)
	'This function is used to identity a valid read-only domain controller.
	On Error Resume Next
	Const ADS_SCOPE_SUBTREE = 2
	Dim objDC,objProperty
	Set objDC = GetObject("LDAP://CN=" & DCName & ",OU=Domain Controllers," & strDefaultNamingContext)
	If Err.Number = -2147016656 Then
		TestOSCIsRODC = Null
	End If
	Call objDC.GetInfoEx(Array("msDS-isRODC"),0)
	TestOSCIsRODC = objDC.[msDS-isRODC]
End Function

Function GetOSCADObjectDN(SamAccountName,ObjectCategory)
	'This function is used to get distinguishedName of an AD object.
	Dim objConnection,objCommand,objRecordSet,strDN
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
	"<LDAP://" & strDefaultNamingContext & _
	">;(&(objectCategory=" & ObjectCategory & ")(sAMAccountName=" & SamAccountName & _
	"));distinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	If objRecordSet.RecordCount = 1 Then
		While Not objRecordSet.EOF
			GetOSCADObjectDN = objRecordSet.Fields("distinguishedName")
			objRecordSet.MoveNext
		Wend
	Else
		GetOSCADObjectDN = Null
	End If
	objConnection.Close
End Function

Sub GetOSCADGroupMemberDN(GroupDN)
	'This function is used to get all group members distinguishedName in one group.
	'Nested group is supported.
	Dim objADGroup,objADGroupMember,strDN
	Set objADGroup = GetObject("LDAP://" & GroupDN)
	For Each objADGroupMember In objADGroup.Members
		If objADGroupMember.Class = "group" Then
			GetOSCADGroupMemberDN(objADGroupMember.distinguishedName)
		Else
			ReDim Preserve arrGroupMemberDNs(i)
			arrGroupMemberDNs(i)= objADGroupMember.distinguishedName	
			i = i + 1
		End If
	Next
End Sub

Function InvokeOSCRODCPwdCachePrepopulation(GroupName,RODCName,WritableDCName)
	'This function is used to execute repadmin command.
	Dim strGroupDN,strGroupMemberDN,strCommand
	'Call GetOSCADObjectDN function to get the distinguishedName of a group object.
	strGroupDN = GetOSCADObjectDN(GroupName,"group")
	'Call GetOSCADGroupMemberDN function to get the distinguishedName of group members in a specified group.
	Call GetOSCADGroupMemberDN(strGroupDN)
	For Each strGroupMemberDN In arrGroupMemberDNs
		'Create commands
		strCommand = "repadmin /rodcpwdrepl " & RODCName & " " & WritableDCName & " """ & strGroupMemberDN & """"
		Set objExec = objShell.Exec(strCommand)
		Do While (objExec.Status <> WshRunning)
			Call WScript.Sleep(200)
		Loop
		strOutput = objExec.StdOut.ReadAll()
		WScript.Echo strOutput			
	Next
End Function

Sub OSCScriptUsage
	WScript.Echo "How to use this script:" & vbCrLf & vbCrLf _
	& "1. Logon to one Writable Domain Controller." & vbCrLf _
	& "2. Open an elevated command console." & vbCrLf _
	& "3. Run following command:" & vbCrLf  _
	& "cscript //nologo PrepopulatePasswordCacheForRODC.vbs " _
	& "/GroupName:""GroupName"" /RODCName:""RODCName"" /WritableDCName:""WritableDCName"" > result.txt"
End Sub

Sub Main
	Dim strGroupName,strRODCName,strWritableDCName,strGroupDN,objArgs,i
	Dim objRegExp
	Set objRegExp = New RegExp
	Set objArgs = WScript.Arguments
	objRegExp.Global = True
	objRegExp.IgnoreCase = True
	objRegExp.Pattern = "groupname|rodcname|writabledcname"
	'Please run this script with cscript.
	If InStr(WScript.FullName,"cscript") = 0 Then
		Call OSCScriptUsage
		WScript.Quit(1)
	End If
	'You must run this script in an elevated command consle
	'and with Domain Administrator privilege.
	If Not TestOSCUserPrivilege Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	End If
	'Verify Arguments
	If WScript.Arguments.Named.Count <> 3 Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	Else
		For i = 0 To objArgs.Count - 1
			If Not objRegExp.Test(objArgs(i)) Then
				Call OSCScriptUsage()
				WScript.Quit(1)
			Else
				With objArgs.Named
					If .Exists("groupname") Then strGroupName = .Item("groupname")
					If .Exists("rodcname") Then strRODCName = .Item("rodcname")
					If .Exists("writabledcname") Then strWritableDCName = .Item("writabledcname")
				End With
			End If
		Next
	End If
	strGroupDN = GetOSCADObjectDN(strGroupName,"group")
	If IsNull(strGroupDN) Then
		WScript.Echo "Please enter a valid group name."
		WScript.Quit(1)
	End If
	If Not TestOSCIsRODC(strRODCName) Or IsNull(TestOSCIsRODC(strRODCName)) Then
		WScript.Echo strRODCName & " is not a read-only Domain Controller."
		WScript.Quit(1)
	End if		
	If TestOSCIsRODC(strWritableDCName)Or IsNull(TestOSCIsRODC(strWritableDCName)) Then
		WScript.Echo strWritableDCName & " is not a Writable Domain Controller."
		WScript.Quit(1)
	End If
	WScript.Echo "Starts at " & Now
	Call InvokeOSCRODCPwdCachePrepopulation(strGroupName,strRODCName,strWritableDCName)
End Sub

Call Main()
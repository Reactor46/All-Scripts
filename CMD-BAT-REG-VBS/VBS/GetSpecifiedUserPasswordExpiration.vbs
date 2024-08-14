'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 


'Run script with administrator privilege
If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "open", 1
Else
	Set objshell  = CreateObject("wscript.shell")
	'Get current user name 
	userName = objshell.ExpandEnvironmentStrings("%UserName%")
	'Get input user name
	WScript.StdOut.Write("Enter the SamAccountName:")
	samaccountName = WScript.StdIn.ReadLine()
	'If input value is null, samaccountname equals username
	If samaccountName = "" Then 
		samaccountName = userName
		strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=*" & samaccountName & "*))"
	'If input value is "*", it will return all records
	ElseIf samaccountName = "*" Then 
		strFilter = "(&(objectCategory=person)(objectClass=user))"
	Else
		strFilter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=*" & samaccountName & "*))"
	End If 
	'create ADO connection
	Set ADOConnection = CreateObject("ADODB.Connection")
	Set ADOCommand = CreateObject("ADODB.Command")
	ADOConnection.Provider = "ADsDSOOBject"
	ADOConnection.Open "Active Directory Provider"
	Set ADOCommand.ActiveConnection = ADOConnection
	' Determine the DNS domain from the RootDSE object.
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDomainDN = objRootDSE.Get("DefaultNamingContext")
	Const ADS_UF_PASSWD_CANT_CHANGE = &H40
	Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
	'Get max password expired days 
	Set oDomain = GetObject("LDAP://" & strDomainDN)
	Set maxPwdAge = oDomain.Get("maxPwdAge")
	numDays = CCur((maxPwdAge.HighPart * 2 ^ 32) + maxPwdAge.LowPart) / CCur(-864000000000)
	
	ADOCommand.CommandText = "<LDAP://" & strDomainDN & ">;" & strFilter & ";distinguishedName,pwdLastSet,userAccountControl;subtree"
	ADOCommand.Properties("Page Size") = 1000
	Set ADORecordset = ADOCommand.Execute
	Do Until ADORecordset.EOF
		strDN = ADORecordset.Fields("distinguishedName").Value
	    regFlag = ADORecordset.Fields("userAccountControl").Value
	    Flag  = True
	    
	    If ((regFlag And ADS_UF_PASSWD_CANT_CHANGE) <> 0) Then
	    	Flag = False
	        PassowrdExpired = "Never"
	    End If
	    If ((regFlag And ADS_UF_DONT_EXPIRE_PASSWD) <> 0) Then
	        Flag = False
	        PassowrdExpired = "Never"
	    End If
	    If (TypeName(ADORecordset.Fields("pwdLastSet").Value) = "Object") Then
		        Set objDate = ADORecordset.Fields("pwdLastSet").Value
		        PasswordLastSetDate = ChangeTimeZone(objDate, reg)
		    Else
		        PasswordLastSetDate = #12/31/1600#
		End If
		If Flag =True  Then 
		    PassowrdExpired = DateAdd("d",numDays,PasswordLastSetDate)
	    End If 
	    strPosition = InStr(strDN,",")
	    samAccountName = Mid(strDN,4,strPosition-4)
	  	wscript.stdout.writeline "UserName           : " & samAccountName 
	  	wscript.stdout.writeline "PassowrdExpiredDate: " & PassowrdExpired 
	  	wscript.stdout.writeline "PasswordLastSetDate: " & PasswordLastSetDate
	  	WScript.stdout.writeline " "
	  	ADORecordset.MoveNext
	Loop 
	
	'This function is to change timezone
	Function ChangeTimeZone(ByVal objDate, ByVal reg)
	    Dim regAdjust, regDate, regHigh, regLow
	    regAdjust = reg
	    regHigh = objDate.HighPart
	    regLow = objdate.LowPart
	    If (regLow < 0) Then
	        regHigh = regHigh + 1
	    End If
	    If (regHigh = 0) And (regLow = 0) Then
	        regAdjust = 0
	    End If
	    regDate = #12/31/1600# + (((regHigh * (2 ^ 32)) + regLow) / 600000000 - regAdjust) / 1440
	    On Error Resume Next
	    ChangeTimeZone = CDate(regDate)
	    If (Err.Number <> 0) Then
	        On Error GoTo 0
	        ChangeTimeZone = #12/31/1600#
	    End If
	    On Error GoTo 0
	End Function
	WScript.StdOut.Write("Press ENTER to exist...")
	qiut = WScript.StdIn.ReadLine()
	WScript.Quit
End If 	
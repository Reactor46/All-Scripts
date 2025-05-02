'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 31/01/2008
' Updated : 09/09/2008
'=*=*=*=*=*=*=*=*=*=*=*=
'*================================================*
'* This Script Finds the Last Logon time of all the Users Listes in
'* a Text File (OpenFile) and there Disabled State
'* The Script Outputs the Returned Value to a Log File (strLogFile)
'*================================================*
'================================================
' Consts
'================================================
'File Consts
Const ForReading = 1
Const ForAppending = 8
'File to be Opend
OpenFile = fOpenFile
'Log file path
Const StrLogFile = "C:\UserMgr1.csv"

'AD Const - Account Disabled
Const ADS_UF_ACCOUNTDISABLE = &H0002
 
'================================================
' Dims and Variables
'================================================
Dim oUser
Dim iResult
Dim UserDisabled
Dim intUAC
Dim intLastLogon
Dim objLastLogon

'================================================
' Sets
'================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")

'==========
'Functions
'==========
Function Log(sUser,sDate,Disabled,Hidden,ExDN)
' This Function Logs the Return Values of the Script
' 	Function Recieves : User Name, Last Logon Date, Disabled State
' 	Function Returns : Writes to the Log File
  If Not objFSO.FileExists(StrLogFile) Then
    Set oTS = objFSO.OpenTextFile(strLogFile, ForAppending, True)
    strHdrs = "Account" & "," & "Inactive From" & "," & "Disabled" & "," & "Mailbox"
    oTS.WriteLine strHdrs
  Else
    Set oTS = objFSO.OpenTextFile(strLogFile, ForAppending, True)
  End If
  strMsg = sUser & "," & sDate

  ' Check if the User is Disabled	
  If Disabled = "True" Then
    strMsg = strMsg & "," & "Disabled"
  Else
    strMsg =strMsg & ","
  End If

  ' Check if the User's Mailbox is Hidden or has no Mailbox
  If Hidden = "True" Then
    strMsg = strMsg & "," & "Hidden"
  ElseIf InStr(ExDN,"/")<1 Then
    strMsg = strMsg & "," & "No Mailbox"
  End If

  ' Write all to the Log
  oTS.WriteLine(strMsg)	

  ' Close the Log File Each Time to Deny File Lock
  oTS.Close
  Set oTS = Nothing
End Function

Function FindUser(strUser)
' This Function Searches the AD for the ADSPath of the User
' 	Function Recieves : User Name to Find
' 	Function Returns : The Users ADsPath if Found and 0 if no Object Found
	Const ADS_SCOPE_SUBTREE = 2
	Dim objRootDSE,objConnection,objCommand,objRecordSet
	Dim strDomainLdap

	Set objRootDSE = GetObject ("LDAP://rootDSE")
	strDomainLdap  = objRootDSE.Get("defaultNamingContext")
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.CommandText = _
		"SELECT AdsPath FROM 'LDAP://" & strDomainLdap & "' WHERE objectClass='user' and sAMAccountName='" &_
			strUser & "'"
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Timeout") = 30
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False
	
	Set objRecordSet = objCommand.Execute
	
	If objRecordSet.RecordCount = 0 Then 
		FindUser = 0
	Else 
		objRecordSet.Requery
		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			FindUser = objRecordSet.Fields("AdsPath").Value
			objRecordSet.MoveNext
		Loop
	End If 	
	objRecordSet.Close
End Function 

Function GetLastLogon(strUser)
' This Function Connects to the User Information and Gets the Last Logon Date
' 	Function Recieves : User Name
' 	Function Returns : Last Logon Date
On error resume Next
' Some Users Doesnt have this information, On Error Resume Next Return Nothing
	Set oUser = GetObject(FindUser(strUser))
	oUser.GetInfo
	' Get The Users Last Logon Time stamp
	Set objLastLogon = oUser.Get("LastLogonTimestamp")
	' Calculate the Users Last Logon High Part and Low Part (the Time Stamp is a Long Integer)
	GetLastLogon = objLastLogon.HighPart * (2^32) + objLastLogon.LowPart
	
End Function

Function fOpenFile
  ' Opening File
  Dim FileLoc

  Set objDialog = CreateObject("UserAccounts.CommonDialog")
  FileLoc = ""

  objDialog.Filter = "Text Files|*.txt|CSV Files|*.csv"
  objDialog.FilterIndex = 1
  objDialog.InitialDir = "C:\"
  intResult = objDialog.ShowOpen
   
  If intResult = 0 Then
      Wscript.Quit
  Else
      FileLoc = objDialog.FileName
  End If

  fOpenFile = FileLoc
End Function

'==========
'Main Code
'==========
' Open the OpenFile Const and Read the User Names
Set objTextFile = objFSO.OpenTextFile(OpenFile, ForReading)

Do Until objTextFile.AtEndOfStream
	intLastLogon = ""
	' Read the File and Calculate the Last Logon and Format it to Day Time
	intLastLogon = GetLastLogon(objTextFile.ReadLine)
	intLastLogon = intLastLogon / (60 * 10000000)
	intLastLogon = intLastLogon / 1440
	' Init Boolean vars as False
	UserDisabled = "False"
	UserHidden = "False"
	
	' Get User Properties
	On Error Resume Next
	intUAC = oUser.Get("userAccountControl")
	blnHidden = oUser.Get("msExchHideFromAddressLists")
	
	' Check the Users Exchange Legacy DN if Exists - if not then no Mailbox - Write to log
	UserExDN = oUser.Get("legacyExchangeDN")
	
	' Check the User Account Control for User Disabled State
	If intUAC And ADS_UF_ACCOUNTDISABLE Then
			UserDisabled = "True"
	Else
			UserDisabled = "False"
	End If
	
	' Check if the Users Mailbox is Hidden from Address Lists
	If blnHidden Then
		UserHidden = "True"
	Else
		UserHidden = "False"
	End If
	
	
	' Users Name
	sName = oUser.sAMAccountName
	
	' Log the Results : the User Name (sName), the Last Logon Date and Format it with #1/1/1601# and the User Disabled State
	 iResult = Log(sName,intLastLogon + #1/1/1601#,UserDisabled,UserHidden,UserExDN)
Loop

' Close the File that was Read
objTextFile.Close
Set objFSO = Nothing
MsgBox "All Done!" & vbNewLine & "File Saved in " & StrLogFile
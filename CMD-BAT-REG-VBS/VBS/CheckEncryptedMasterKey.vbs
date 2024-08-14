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
'---------------------
' String: en-US
'---------------------
Const INPUT_INSTANCE_NAME = "Please input your instance name. (ServerName\InstanceName)"
Const SQL_INSTANCE = "Create SQL Connection"
Const INSTANCE_FAILED = "Fail to create SqlConnection object!"
Const LIST_OBJECTS = "List the database master key which is not encrypted by service master key."
Const No_ENCRYPTED_MASTER = "Detected your database master key has been encrypted by service master key or no TDE-enabled database exists."


Dim errFlag,connString,oConnection,outPut
errFlag = False
outPut = ""


'---------------------
' Main
'---------------------
Call Main()


'Main logic
Sub Main()
	'Create SQLConnection object
	Call NewOSCSQLConnection()
	
	If IsEmpty(oConnection) Then
		Exit Sub
	End If
	
	'Check not encrypted master key
	Call GetOSCNotEncryptedMasterKey()
	
	'Output Result
	If errFlag = False Then
		If outPut <> ""Then
			WScript.Echo LIST_OBJECTS & _
						 vbCrLf & _
						 vbCrLf & _
						 outPut
		Else
			WScript.Echo No_ENCRYPTED_MASTER
		End If
	End If
End Sub

'Create a new SqlConnection Object
Sub NewOSCSQLConnection()
	Dim sqlInstance
	sqlInstance = InputBox(INPUT_INSTANCE_NAME,SQL_INSTANCE)
	
	If sqlInstance = "" Then
		WScript.Echo INSTANCE_FAILED
		Exit Sub
	End If
	
	connString = "Provider=SQLOLEDB;Data Source=" & _
				 sqlInstance & _
				 "; Initial Catalog=master;Integrated Security=SSPI"
	Set oConnection = CreateObject("ADODB.Connection")
End Sub

'Get database master key which is not encrypted by service master key.
Sub GetOSCNotEncryptedMasterKey()
	Dim sqlQuery,oResults,conError
	sqlQuery = "SELECT A.IsMasterKeyEncryptedByServer,B.CountOfDBEnabledForTDE FROM (SELECT CASE WHEN is_master_key_encrypted_by_server = 1 THEN 'TRUE' ELSE 'FALSE' END AS IsMasterKeyEncryptedByServer FROM sys.databases WHERE  name = 'master') AS A CROSS JOIN (SELECT Count(is_encrypted) AS CountOfDBEnabledForTDE FROM sys.databases WHERE is_encrypted = 1) AS B"
	outPut = ""
	
	'1: Open | 0: Close
	If oConnection.State <> 1 Then
		On Error Resume Next
		Err.Clear
		oConnection.Open connString
		
		If Err.Number <> 0 Then
			conError = "Error: " & Err.Number & vbCrLf & _
					   "Error(Hex): " & Hex(Err.Number) & vbCrLf & _
					   "Source: " & Err.Source & vbCrLf & _
					   "Description: " & Err.Description & vbCrLf
			
			WScript.Echo conError
			errFlag = True
			Err.Clear()
			Exit Sub
		End If
		
		Set oResults = oConnection.Execute(sqlQuery)
		'State = 1 indicates that the recordset is open.
		If oResults.State = 1 Then
			Do While Not oResults.EOF
				If oResults("IsMasterKeyEncryptedByServer") = "TRUE" Or _
				   oResults("CountOfDBEnabledForTDE") = 0 Then
					outPut = ""
					Exit Sub
				End If
				
				outPut = outPut & "[IsMasterKeyEncryptedByServer: " & oResults("IsMasterKeyEncryptedByServer") & _
						 " , CountOfDBEnabledForTDE: " & oResults("CountOfDBEnabledForTDE") & "];" & _
						 vbCrLf
				
				oResults.MoveNext
			Loop
		End If
		
		oResults.Close()
		oConnection.Close()	
	End If	
End Sub




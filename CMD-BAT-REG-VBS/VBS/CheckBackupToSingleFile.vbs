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
Const LIST_OBJECTS = "List the objects whose backups appended to existing backup devices."
Const NO_RECORDS_FOUND = "Cannot find backups appended to existing backup devices."


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
	
	'Check backups appended to existing backup devices 
	Call GetOCSBackupToSingleFile()
	
	'Output Result
	If errFlag = False Then
		If outPut <> ""Then
			WScript.Echo LIST_OBJECTS & _
						 vbCrLf & _
						 vbCrLf & _
						 outPut
		Else
			WScript.Echo NO_RECORDS_FOUND
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

'Get backups appended to existing backup devices
Sub GetOCSBackupToSingleFile()
	Dim sqlQuery,oResults,conError
	sqlQuery = "SELECT m.name,m.media_set_id,m.position FROM (SELECT d.name,media_set_id,position FROM msdb..backupset AS e JOIN (SELECT a.name,b.database_guid,c.family_guid FROM master..sysdatabases AS a JOIN sys.database_recovery_status AS b ON a.dbid = b.database_id JOIN sys.database_recovery_status AS c ON a.dbid = c.database_id) AS d ON e.database_guid = d.database_guid AND e.family_guid = d.family_guid) AS m JOIN (SELECT media_set_id,Max(position) AS position FROM msdb..backupset AS e JOIN (SELECT a.name,b.database_guid,c.family_guid FROM master..sysdatabases AS a JOIN sys.database_recovery_status AS b ON a.dbid = b.database_id JOIN sys.database_recovery_status AS c ON a.dbid = c.database_id) AS d ON e.database_guid = d.database_guid AND e.family_guid = d.family_guid GROUP BY media_set_id,d.name HAVING Max(position) > 1) AS n ON m.media_set_id = n.media_set_id AND m.position = n.position"
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
				outPut = outPut & "[name: " & oResults("name") & _
						 " , media_set_id: " & oResults("media_set_id") & _
						 " , position: " & oResults("position") & "];" & _
						 vbCrLf
				
				oResults.MoveNext
			Loop
		End If
		
		oResults.Close()
		oConnection.Close()	
	End If	
End Sub




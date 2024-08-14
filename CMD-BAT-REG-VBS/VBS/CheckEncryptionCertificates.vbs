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
Const LIST_OBJECTS = "List the objects whose encryption certificates without a backup."
Const NO_RECORDS_FOUND = "Cannot find encryption certificates without a backup."


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
	
	'Check encryption certificates
	Call GetOSCEncryptionCertificates()
	
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

'Get encryption certificates without backup
Sub GetOSCEncryptionCertificates()
	Dim sqlQuery,oResults,conError
	sqlQuery = "SELECT * FROM (SELECT a.name AS Database_Name,COALESCE(c.name, 'NA')AS Certificate_Name, CASE a.is_encrypted WHEN 1 THEN 'Encrypted' WHEN 0 THEN 'Not Encrypted' END AS Is_Encrypted,COALESCE(Cast(c.pvt_key_last_backup_date AS VARCHAR(50)), 'NA')AS Last_BackupDate FROM sys.databases a LEFT OUTER JOIN sys.dm_database_encryption_keys b ON a.database_id = b.database_id LEFT OUTER JOIN sys.certificates c ON b.encryptor_thumbprint = c.thumbprint JOIN sys.database_mirroring d ON d.database_id = a.database_id WHERE d.mirroring_guid IS NULL AND source_database_id IS NULL AND a.name NOT IN ( 'master', 'tempdb', 'model', 'msdb' )) AS EncryptedCertificates WHERE last_backupdate = 'NA' AND certificate_name <> 'NA'"
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
				outPut = outPut & "[Database_Name: " & oResults("Database_Name") & _
						 " , Certificate_Name: " & oResults("Certificate_Name") & _
						 " , Is_Encrypted: " & oResults("Is_Encrypted") & _
						 " , Last_BackupDate: " & oResults("Last_BackupDate") & "];" & _
						 vbCrLf
				
				oResults.MoveNext
			Loop
		End If
		
		oResults.Close()
		oConnection.Close()	
	End If	
End Sub




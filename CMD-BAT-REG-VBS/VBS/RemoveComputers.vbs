'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------

Option Explicit 

Sub main()
	Dim objshell, FSO, Count, FilePath, GroupName
	Set objshell = CreateObject("wscript.shell")
	Set FSO = CreateObject("Scripting.FileSystemobject")
	'Check arguments count 
	Count  =WScript.Arguments.Count 
	Select Case Count 
		Case 2 
			'Get the file 
			FilePath = WScript.Arguments.Item(0)
			'Get group name 
			GroupName = WScript.Arguments.Item(1)
			'Check if the file exists 
			If FSO.FileExists(FilePath) Then 
				Dim objRootDSE, strDomain , objConnection, objRecordSet
				Dim strDN, TextFile, strName, CompAdsPath, objGroup,GroupAdsPath
				Const ADS_SCOPE_SUBTREE = 2
				' Get domain components
				Set objRootDSE = GetObject("LDAP://RootDSE")
				strDomain = objRootDSE.Get("DefaultNamingContext")
				' Get username to search for
				Set objConnection = CreateObject("ADODB.Connection")
				objConnection.Provider = "ADsDSOObject"
				objConnection.Open "Active Directory Provider"
				Dim oRs
				'Check if the group exists
				Set oRs = objConnection.Execute("SELECT adspath FROM 'LDAP://" & strDomain & "'" & "WHERE objectCategory='group' AND " & "Name='" & GroupName & "'")
				If Not oRs.EOF Then
					GroupAdsPath = oRs("adspath")
				End If
				If IsEmpty(GroupAdsPath) = False  Then 
					Set TextFile = FSO.OpenTextFile(FilePath)
					Do Until Textfile.AtEndOfStream 
						strName  = TextFile.ReadLine
						Set objRecordSet = objConnection.Execute("SELECT adspath FROM 'LDAP://" & strDomain & "' WHERE objectCategory='Computer' AND Name = '" & strName & "'")
						' If Computer was found
						If Not objRecordSet.EOF Then
							CompAdsPath = objRecordSet("adspath")
							
							Set objGroup = GetObject(GroupAdsPath) 
							'If the computer is a member of the group 
							If (objGroup.IsMember(CompAdsPath) = True ) Then
							
								'Remove computer from group 
								objGroup.Remove(CompAdsPath)
								WScript.Echo "Remove " &  strName & " successfully." 
							Else
								WScript.Echo "Computer " & strName  & " not in group"
							End If
						Else 
							WScript.Echo  "Computer " & strName  & " not found."
						End If	
						Set objGroup = Nothing 
						Set objRecordSet = Nothing 
						strName = Null	
					Loop 		
				Else 
					WScript.Echo "The specified group not exist."
				End If 
				Set objRootDSE = Nothing 
				objConnection.Close
			Else
				WScript.Echo "The specified file not found."
			End If 
		Case Else 
			WScript.Echo "Invalid input value, please try again."
			
		End Select 
End Sub

Call main 

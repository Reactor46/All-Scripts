'*******************************************************************
'*** Description : This Script extract the permissions of each directories
'***               into a HTML page
'*** Auteur      : F.P.IVART
'*** Version     : 1.0.0
'*** Date modif  : 26/06/2009 by F.P.IVART
'*** Input       : /P: for the directory you need to extract permissions
'***               /L: for the number of subdirectories you want explore
'*** Output      : Fichier HTML : On c:\Temp   directory
'*** Remarque    : 
'*******************************************************************

Option Explicit

Dim objFSO
Dim objFile
Dim objOutputFile
Dim WshShell
Dim LevelArboMax
Dim sOutputFileName
Dim DataListRights
Dim oFolder
Dim Path
Dim indSousRep 
Dim indGroupe
Dim LastRep

Set WshShell = createObject ("Wscript.shell")
Set objFSO = CreateObject ("Scripting.FileSystemObject")

'*** Get arguments 
If WScript.Arguments.named.Count = 0 Then
	Wscript.echo "This script need at least one argument" + Chr(10) + "Use :ACL2HTM.vbs " + Chr(10) + "Options /P:Path /L:Level"
	wscript.quit(1)
End If

'**** Directory to explore
If WScript.Arguments.Named.exists("P") Then
	Path = WScript.Arguments.Named.Item("P")
Else
	Wscript.echo "This script need at least the argument /P " + Chr(10) + "Use :ACL2HTM.vbs " + Chr(10) + "Options /P:Path /L:Level"
	wscript.quit(2)
End If

LastRep = Path
Do While InStr(LastRep, "\") <> 0 
		LastRep = Right(LastRep, Len(LastRep) - InStr(LastRep, "\"))
Loop


'**** Level of sub directories
If WScript.Arguments.Named.exists("L") Then
	LevelArboMax = Int(WScript.Arguments.Named.Item("L"))
	If LevelArboMax < 0 Then
		LevelArboMax = 3
	End if
Else
	LevelArboMax = 3
End If

'*** Output HTML file
sOutputFileName = "C:\temp\PermissionsDirectories_" & LastRep & ".htm"
Set objFile = objFSO.CreateTextFile(sOutputFileName)
Set objFile = Nothing
Set objOutputFile = objFSO.OpenTextFile(sOutputFileName, 2, True)

'*** Header of HTML flie 
objOutputFile.WriteLine "<html>"
objOutputFile.WriteLine "<TABLE style=""FONT-SIZE: 10pt"" face=""Times New Roman"" BGCOLOR=""#FFFFFF"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">"
objOutputFile.WriteLine "<TR ALIGN=""CENTER"" style=""FONT-SIZE: 12pt"" face=""Times New Roman"">LIST OF PERMISSIONS AND GROUP FOR DIRECTORIES " & Path & "<BR><BR></TR>"
objOutputFile.WriteLine "<TR style=""FONT-SIZE: 8pt"">Last Update : " & Date & "<BR><BR></TR>"
objOutputFile.WriteLine "<TR BGCOLOR=""#9999ff"">"
objOutputFile.WriteLine "<TD></TD><TD> Parent DIR</TD>"
If LevelArboMax > 0 Then
	For indSousRep = 1 To LevelArboMax
			objOutputFile.WriteLine "<TD> Child DIR " & indSousRep & "</TD>"
	Next
End If

'*** 10 permissions seem's to me the Maxmum for header but the tab could be most
For indGroupe = 1 To 10 
	objOutputFile.WriteLine "<TD> Group " & indGroupe & "</TD>"
Next

objOutputFile.WriteLine "</TR>"

'*** Initialisation list of permissions for the directory
Set DataListRights = CreateObject("System.Collections.ArrayList")

Set oFolder = objFSO.getFolder(Path)

'*** Walk into subdirectories in order to find permissions
ParcourRepertoire oFolder, LevelArboMax


objOutputFile.WriteLine "</TABLE></html>"

objOutputFile.Close
		
Set oFolder = Nothing
Set DataListRights = Nothing
Set objOutputFile = Nothing
Set objFSO = Nothing
	
Wscript.QUIT


'*******************************************************************
'*** Description : Walk into subdirectories 
'*** Auteur      : F.P.IVART
'*** Version     : 1.0.0
'*** Date modif  : 26/06/2009 by F.P.IVART
'*** Input       : - Path : Directory you want to extract permissions
'***               - Level  : Level of subdirectories
'*** Output      : 
'*** Remarque    : Carreful : It's a recurssive function
'*******************************************************************
Function ParcourRepertoire(Path, Level)
	
	Dim SousRepertoires
	Dim UnSousRep

	SeePermissions Path.path, Level

	If Len(Path.path) < 254 And Level > 0 Then
		Set SousRepertoires = Path.SubFolders
		If SousRepertoires.Count <> 0 Then
			For Each UnSousRep In SousRepertoires
				ParcourRepertoire UnSousRep, Level - 1
			Next
		End If

		If err.number<>0 Then
			wscript.echo "Error in directory " & Path.Path & Chr(10) & err.description
			Err.clear
		End if
	End If

End Function

'*******************************************************************
'*** Description : See the directory permissions
'*** Auteur      : F.P.IVART
'*** Version     : 1.0.0
'*** Date modif  : 26/06/2009 by F.P.IVART
'*** Input       : - Path : directory you want to extract permissions
'***               - Level  : Level of subdirectories
'*** Output      : 
'*** Remarque    : Carreful : It's a recursive function
'*******************************************************************
Function SeePermissions(Path, Level)

	Dim oAce '*** variable for the new ACE
	Dim oSD  '*** variable for the Security Descriptor of the object
	Dim oDacl '*** variable for the DACL of the object
	Dim oADsSecurityUtility 
	Dim wmiAce
	'**** Dim Proprietaire
	Dim DataListRightsEnCours 
	Dim Arboresence
	Dim sLigne
	Dim Rights
	Dim Groupe
	Dim bTrouve
	Dim Apres 
	Dim Avant 
	Dim nbRep 
	Dim strItem, strItem1, strItem2
	
	Set DataListRightsEnCours = CreateObject("System.Collections.ArrayList")
	Set oADsSecurityUtility = CreateObject("ADsSecurityUtility")
	Set oSD = oADsSecurityUtility.GetSecurityDescriptor(Path, 1, 1)

	If Err.number <> 0 Then
		wscript.echo "Erreur sur le répertoire "& Path
		Err.clear
	End If

	'*** Proprietaire = Osd.Owner

	Set oDacl = oSD.DiscretionaryAcl
	
	sLigne = ""
	Apres = ""
	Avant = ""

	'*** Add the ";" before the currently directory
	'*** Evite d'afficher l'arboresence complète 	
	For nbRep = 1 To  Level
			Apres = Apres + "</TD><TD>"
	Next 
			
	Arboresence = Path
	Do While InStr(Arboresence, "\") <> 0 
			Avant  = Avant + "</TD><TD>"
			Arboresence = Right(Arboresence, Len(Arboresence) - InStr(Arboresence, "\"))
	Loop

	sLigne =  Avant + Arboresence + Apres 
	
	'*** We follow the ACL permissions
	For Each wmiAce in oDACL
		Select Case int(wmiAce.AccessMask)
			Case 2032127
				Rights = "FULL"
			Case 1179817
				Rights = "RX"
			Case -1610612736
				Rights = "RXe"
			Case 1245631
				Rights = "RWX"
			Case 268435456
				Rights = "FULL SUB ONLY"
			Case else
				Rights = Cstr(wmiAce.AccessMask)
		End Select

		'*** If you want to make a restriction only with directories who start like "NE-"
		'If InStr(wmiAce.Trustee,"NE-") Then
			Groupe = Right(wmiAce.Trustee, Len(wmiAce.Trustee) - InStr(wmiAce.Trustee,"\"))
			
			'*** We feeling the permissions list
			DataListRightsEnCours.Add Groupe & " (" & Rights & ")" 
			'DataListRightsEnCours.Add Groupe

			'*** We verrify that the permission isn't allready in the last list
			bTrouve = False
			For Each strItem in DataListRights
					If strItem = Groupe & " (" & Rights & ")" Then 
						bTrouve = True
						Exit For
					End If
			Next

			If Not bTrouve Then
					'*** If the permission is not present in the last list we addition it
					DataListRights.Add Groupe & " (" & Rights & ")" 
					'DataListRights.Add Groupe
			End If
		'End If
	Next
	
	'DataList.Sort() '*** Unfortunatelly we can easy sort the list cause the permission and order isn't the same for each directory
	
	'*** We delete into the old tab the permission who aren't in the new tab
	'*** in order to keep as possible the same order of permissions
	On Error Resume Next '*** The deletion cause the decrease of the original tab
	For Each strItem1 in DataListRights
			bTrouve = False
			For Each strItem2 in DataListRightsEnCours
				If strItem1 = strItem2 Then
					bTrouve = True
					Exit For
				End If
			Next
			If Not bTrouve Then
					'*** If we don't find the permission in the last list we delete them
					'*** This is this permission who will be print in order to keep the sort of print
					DataListRights.Remove(strItem1)
			End If
	Next
	On Error GoTo 0
	
	For Each strItem in DataListRights
			sLigne = sLigne & "</TD><TD>" & strItem 
	Next

	'*** Write the line in the HTML output file
	objOutputFile.WriteLine "<TR BGCOLOR=""#CCCCFF""><TD>" & sLigne & "</TD></TR>"

	Set DataListRightsEnCours = Nothing
	Set oDacl = Nothing
	Set oSD = Nothing
	Set oADsSecurityUtility = Nothing

End Function

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
Dim objFichierSortie
Dim WshShell
Dim NiveauArboMaxi
Dim sNomFichierSortie
Dim DataListDroits
Dim oFolder
Dim Chemin
Dim indSousRep 
Dim indGroupe
Dim DernierRep

Set WshShell = createObject ("Wscript.shell")
Set objFSO = CreateObject ("Scripting.FileSystemObject")

'*** Get arguments 
If WScript.Arguments.named.Count = 0 Then
	Wscript.echo "This script need at least one argument" + Chr(10) + "Use :ACL2HTM_eng.vbs " + Chr(10) + "Options /P:Path /L:Level"
	wscript.quit(1)
End If

'**** Directory to explore
If WScript.Arguments.Named.exists("P") Then
	Chemin = WScript.Arguments.Named.Item("P")
Else
	Wscript.echo "This script need at least the argument /P " + Chr(10) + "Use :ACL2HTM_eng.vbs " + Chr(10) + "Options /P:Path /L:Level"
	wscript.quit(2)
End If

DernierRep = Chemin
Do While InStr(DernierRep, "\") <> 0 
		DernierRep = Right(DernierRep, Len(DernierRep) - InStr(DernierRep, "\"))
Loop


'**** Level of sub directories
If WScript.Arguments.Named.exists("L") Then
	NiveauArboMaxi = Int(WScript.Arguments.Named.Item("L"))
	If NiveauArboMaxi < 0 Then
		NiveauArboMaxi = 3
	End if
Else
	NiveauArboMaxi = 3
End If

'*** Output HTML file
sNomFichierSortie = "C:\temp\PermissionsDirectories_" & DernierRep & ".htm"
Set objFile = objFSO.CreateTextFile(sNomFichierSortie)
Set objFile = Nothing
Set objFichierSortie = objFSO.OpenTextFile(sNomFichierSortie, 2, True)

'*** Header of HTML flie 
objFichierSortie.WriteLine "<html>"
objFichierSortie.WriteLine "<TABLE style=""FONT-SIZE: 10pt"" face=""Times New Roman"" BGCOLOR=""#FFFFFF"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">"
objFichierSortie.WriteLine "<TR ALIGN=""CENTER"" style=""FONT-SIZE: 12pt"" face=""Times New Roman"">LIST OF PERMISSIONS AND GROUP FOR DIRECTORIES " & Chemin & "<BR><BR></TR>"
objFichierSortie.WriteLine "<TR style=""FONT-SIZE: 8pt"">Last Update : " & Date & "<BR><BR></TR>"
objFichierSortie.WriteLine "<TR BGCOLOR=""#9999ff"">"
objFichierSortie.WriteLine "<TD></TD><TD> Parent DIR</TD>"
If NiveauArboMaxi > 0 Then
	For indSousRep = 1 To NiveauArboMaxi
			objFichierSortie.WriteLine "<TD> Child DIR " & indSousRep & "</TD>"
	Next
End If

'*** 10 permissions seem's to me the maximum for header but the tab could be most
For indGroupe = 1 To 10 
	objFichierSortie.WriteLine "<TD> Group " & indGroupe & "</TD>"
Next

objFichierSortie.WriteLine "</TR>"

'*** Initialisation list of permissions for the directory
Set DataListDroits = CreateObject("System.Collections.ArrayList")

Set oFolder = objFSO.getFolder(Chemin)

'*** Walk into subdirectories in order to find permissions
ParcourRepertoire oFolder, NiveauArboMaxi


objFichierSortie.WriteLine "</TABLE></html>"

objFichierSortie.Close
		
Set oFolder = Nothing
Set DataListDroits = Nothing
Set objFichierSortie = Nothing
Set objFSO = Nothing
	
Wscript.QUIT


'*******************************************************************
'*** Description : Walk into subdirectories 
'*** Auteur      : F.P.IVART
'*** Version     : 1.0.0
'*** Date modif  : 26/06/2009 by F.P.IVART
'*** Input       : - Chemin : Directory you want to extract permissions
'***               - Level  : Level of subdirectories
'*** Output      : 
'*** Remarque    : Carreful : It's a recurssive function
'*******************************************************************
Function ParcourRepertoire(Chemin, Level)
	
	Dim SousRepertoires
	Dim UnSousRep

	VoirPermissions Chemin.path, Level

	If Len(Chemin.path) < 254 And Level > 0 Then
		Set SousRepertoires = Chemin.SubFolders
		If SousRepertoires.Count <> 0 Then
			For Each UnSousRep In SousRepertoires
				ParcourRepertoire UnSousRep, Level - 1
			Next
		End If

		If err.number<>0 Then
			wscript.echo "Erreur dans le répertoire " & Chemin.Path & Chr(10) & err.description
			Err.clear
		End if
	End If

End Function

'*******************************************************************
'*** Description : See the directory permissions
'*** Auteur      : F.P.IVART
'*** Version     : 1.0.0
'*** Date modif  : 26/06/2009 by F.P.IVART
'*** Input       : - Chemin : directory you want to extract permissions
'***               - Level  : Level of subdirectories
'*** Output      : 
'*** Remarque    : Carreful : It's a recursive function
'*******************************************************************
Function VoirPermissions(Chemin, Level)

	Dim oAce '*** variable for the new ACE
	Dim oSD  '*** variable for the Security Descriptor of the object
	Dim oDacl '*** variable for the DACL of the object
	Dim oADsSecurityUtility 
	Dim wmiAce
	'**** Dim Proprietaire
	Dim DataListDroitsEnCours 
	Dim Arboresence
	Dim sLigne
	Dim Droits
	Dim Groupe
	Dim bTrouve
	Dim Apres 
	Dim Avant 
	Dim nbRep 
	Dim strItem, strItem1, strItem2
	
	Set DataListDroitsEnCours = CreateObject("System.Collections.ArrayList")
	Set oADsSecurityUtility = CreateObject("ADsSecurityUtility")
	Set oSD = oADsSecurityUtility.GetSecurityDescriptor(Chemin, 1, 1)

	If Err.number <> 0 Then
		wscript.echo "Erreur sur le répertoire "& Chemin
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
			
	Arboresence = Chemin
	Do While InStr(Arboresence, "\") <> 0 
			Avant  = Avant + "</TD><TD>"
			Arboresence = Right(Arboresence, Len(Arboresence) - InStr(Arboresence, "\"))
	Loop

	sLigne =  Avant + Arboresence + Apres 
	
	'*** We follow the ACL permissions
	For Each wmiAce in oDACL
		Select Case int(wmiAce.AccessMask)
			Case 2032127
				Droits = "FULL"
			Case 1179817
				Droits = "RX"
			Case -1610612736
				Droits = "RXe"
			Case 1245631
				Droits = "RWX"
			Case 268435456
				Droits = "FULL SUB ONLY"
			Case else
				Droits = Cstr(wmiAce.AccessMask)
		End Select

		'*** If you want to make a restriction only with directories who start like "NE-"
		'If InStr(wmiAce.Trustee,"NE-") Then
			Groupe = Right(wmiAce.Trustee, Len(wmiAce.Trustee) - InStr(wmiAce.Trustee,"\"))
			
			'*** We feeling the permissions list
			DataListDroitsEnCours.Add Groupe & " (" & Droits & ")" 
			'DataListDroitsEnCours.Add Groupe

			'*** We verrify that the permission isn't allready in the last list
			bTrouve = False
			For Each strItem in DataListDroits
					If strItem = Groupe & " (" & Droits & ")" Then 
						bTrouve = True
						Exit For
					End If
			Next

			If Not bTrouve Then
					'*** If the permission is not present in the last list we addition it
					DataListDroits.Add Groupe & " (" & Droits & ")" 
					'DataListDroits.Add Groupe
			End If
		'End If
	Next
	
	'DataList.Sort() '*** Unfortunatelly we can easy sort the list cause the permission and order isn't the same for each directory
	
	'*** We delete into the old tab the permission who aren't in the new tab
	'*** in order to keep as possible the same order of permissions
	On Error Resume Next '*** The deletion cause the decrease of the original tab
	For Each strItem1 in DataListDroits
			bTrouve = False
			For Each strItem2 in DataListDroitsEnCours
				If strItem1 = strItem2 Then
					bTrouve = True
					Exit For
				End If
			Next
			If Not bTrouve Then
					'*** If we don't find the permission in the last list we delete them
					'*** This is this permission who will be print in order to keep the sort of print
					DataListDroits.Remove(strItem1)
			End If
	Next
	On Error GoTo 0
	
	For Each strItem in DataListDroits
			sLigne = sLigne & "</TD><TD>" & strItem 
	Next

	'*** Write the line in the HTML output file
	objFichierSortie.WriteLine "<TR BGCOLOR=""#CCCCFF""><TD>" & sLigne & "</TD></TR>"

	Set DataListDroitsEnCours = Nothing
	Set oDacl = Nothing
	Set oSD = Nothing
	Set oADsSecurityUtility = Nothing

End Function

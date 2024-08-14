'by K8l0
On Error Resume Next
Const ForAppending = 8
'Commom Permissions
Const FOLDER_FULL_CONTROL = 2032127
Const FOLDER_MODIFY = 1245631
Const FOLDER_READ_ONLY = 1179785
Const FOLDER_READ_CONTENT_EXECUTE =  1179817 
Const FOLDER_READ_CONTENT_EXECUTE_WRITE =  1180095 
Const FOLDER_WRITE = 1179926
Const FOLDER_READ_WRITE = 1180063 
'Special Permissions
Const FOLDER_LIST_DIRECTORY = 1
Const FOLDER_ADD_FILE = 2
Const FOLDER_ADD_SUBDIRECTORY = 4
Const FOLDER_READ_EA = 8
Const FOLDER_WRITE_EA = 16
Const FOLDER_EXECUTE = 32
Const FOLDER_DELETE_CHILD = 64
Const FOLDER_READ_ATTRIBUTES = 128
Const FOLDER_WRITE_ATTRIBUTES = 256
Const FOLDER_DELETE = 65536
Const FOLDER_READ_CONTROL = 131072
Const FOLDER_WRITE_DAC = 262144
Const FOLDER_WRITE_OWNER = 524288
Const FOLDER_SYNCHRONIZE = 1048576
'INHERIT
'Const FOLDER_OBJECT_INHERIT_ACE = 1
'Const FOLDER_CONTAINER_INHERIT_ACE = 2
'Const FOLDER_NO_PROPAGATE_INHERIT_ACE = 4
'Const FOLDER_INHERIT_ONLY_ACE = 8
Const FOLDER_INHERITED_ACE = 16
'ACL Control
Const SE_DACL_PRESENT = 4
Const ACCESS_ALLOWED_ACE_TYPE = 0
Const ACCESS_DENIED_ACE_TYPE  = 1

strComputer = "."

If WScript.Arguments.Count = 3 Then
	strTargetPath = WScript.Arguments.Item(0)
	strOutFile = WScript.Arguments.Item(1)
	strdrop = WScript.Arguments.Item(2)
Elseif WScript.Arguments.Count = 2 Then
	strTargetPath = WScript.Arguments.Item(0)
	strOutFile = WScript.Arguments.Item(1)
	strdrop = ""
Else
	wscript.echo "Run at CMD Prompt: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt"
	wscript.echo "To drop Inheritance run: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt /dropInherit"
	wscript.quit
End If

If Trim(strTargetPath) = "" or Trim(strOutFile) = "" Then
	wscript.echo "Run at CMD Prompt: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt"
	wscript.echo "To drop Inheritance run: cscript List_Sec_Folder_v2.vbs c:\PastaTeste Outlog.txt /dropInherit"
	wscript.quit
End If 

Wscript.echo ""
Wscript.echo "Start Process"
Wscript.echo ""
Wscript.echo "Root Folder Target: " & strTargetPath

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.OpenTextFile(strOutFile, ForAppending, True)
objOutFile.Writeline "Date;Time;Folder;Group / User Name;Commom Permission's;Special Permission's;Access Type;Inheritance;Error's"

ShowSubACL objFSO.GetFolder(strTargetPath)
ShowSubfolders objFSO.GetFolder(strTargetPath)

Wscript.echo ""
Wscript.echo "Finished Process"
Wscript.echo "k8l0"
objOutFile.Close

Sub ShowSubFolders(Folder)
	On Error Resume Next
    For Each Subfolder in Folder.SubFolders
		'Wscript.Echo Subfolder.Path
		'Wscript.echo "Get ACL on Path: " & Subfolder.Path
		ShowSubACL(Subfolder.Path)
		ShowSubFolders Subfolder
		If Err.Number = 0 Then
			strErros = "No Error's"
		ElseIf Err.Number = 451 Then
			strErros = "No Error's"
			Err.clear
		Else
			strErros = "Cod.: " & Err.Number & " Desc.: " & Err.description
			objOutFile.Writeline Date() & ";" & Time() & ";" & FolderPerm & ";" & "" & "" & "" & ";" & "" & ";" & "" & ";" & ""  & ";" & "" & ";" & strErros
			Err.clear
		End If
    Next
End Sub

Sub ShowSubACL(FolderPerm)
	On Error Resume Next
	strCPerm = ""
	strSPerm = ""
	strTypePerm = "" 
	strInherit = ""
	strErros = ""
	Set objWMIService = GetObject("winmgmts:")
	Set objFolderSecuritySettings = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & FolderPerm & "'")
	intRetVal = objFolderSecuritySettings.GetSecurityDescriptor(objSD)
	intControlFlags = objSD.ControlFlags
	If intControlFlags AND SE_DACL_PRESENT Then
		arrACEs = objSD.DACL
		If Err.Number = 0 Then
			strErros = "No Error's"
		Else
			strErros = "Cod.: " & Err.Number & " Desc.: " & Err.description
			objOutFile.Writeline Date() & ";" & Time() & ";""" & FolderPerm  & """;" & "" & "" & "" & ";" & "" & ";" & "" & ";" & ""  & ";" & "" & ";" & strErros
			Err.clear
		End If
		For Each objACE in arrACEs
			
			'ACL Type
			If objACE.AceType = ACCESS_ALLOWED_ACE_TYPE Then strTypePerm = "Allowed"
			If objACE.AceType = ACCESS_DENIED_ACE_TYPE Then strTypePerm = "Denied"
			
			'Inherit
			If objAce.AceFlags AND FOLDER_INHERITED_ACE Then
				strInherit = "Yes"
			Else
				strInherit = "No"
			End if
			
			'Commom Permissions
			If objACE.AccessMask = FOLDER_FULL_CONTROL Then 
				strCPerm = "Full Control"
			ElseIf objACE.AccessMask = FOLDER_MODIFY Then 
				strCPerm = "Modify"
			ElseIf objACE.AccessMask = FOLDER_READ_CONTENT_EXECUTE_WRITE Then
				strCPerm = "Read & Execute / List Folder Contents (folders only) + Write"
			ElseIf objACE.AccessMask = FOLDER_READ_CONTENT_EXECUTE Then 
				strCPerm = "Read & Execute / List Folder Contents (folders only)"
			ElseIf objACE.AccessMask = FOLDER_READ_WRITE Then
				strCPerm = "Read + Write"
			ElseIf objACE.AccessMask = FOLDER_READ_ONLY Then 
				strCPerm = "Read Only"
			ElseIf objACE.AccessMask = FOLDER_WRITE Then 
				strCPerm = "Write"
			Else
				strCPerm = "Special"
			End If
			
			'Special Permissions
			strSPerm = ""
			If objACE.AccessMask and FOLDER_EXECUTE Then strSPerm = strSPerm & "Traverse Folder/Execute File, "
			If objACE.AccessMask and FOLDER_LIST_DIRECTORY Then strSPerm = strSPerm & "List Folder/Read Data, "
			If objACE.AccessMask and FOLDER_READ_ATTRIBUTES Then strSPerm = strSPerm & "Read Attributes, "
			If objACE.AccessMask and FOLDER_READ_EA Then strSPerm = strSPerm & "Read Extended Attributes, "
			If objACE.AccessMask and FOLDER_ADD_FILE Then strSPerm = strSPerm & "Create Files/Write Data, "
			If objACE.AccessMask and FOLDER_ADD_SUBDIRECTORY Then strSPerm = strSPerm & "Create Folders/Append Data"
			If objACE.AccessMask and FOLDER_WRITE_ATTRIBUTES Then strSPerm = strSPerm & "Write Attributes, "
			If objACE.AccessMask and FOLDER_WRITE_EA Then strSPerm = strSPerm & "Write Extended Attributes, "
			If objACE.AccessMask and FOLDER_DELETE_CHILD Then strSPerm = strSPerm & "Delete Subfolders and Files, "
			If objACE.AccessMask and FOLDER_DELETE Then strSPerm = strSPerm & "Delete, "
			If objACE.AccessMask and FOLDER_READ_CONTROL Then strSPerm = strSPerm & "Read Permissions, "
			If objACE.AccessMask and FOLDER_WRITE_DAC Then strSPerm = strSPerm & "Change Permissions, "
			If objACE.AccessMask and FOLDER_WRITE_OWNER Then strSPerm = strSPerm & "Take Ownership, "
			If objACE.AccessMask and FOLDER_SYNCHRONIZE Then strSPerm = strSPerm & "Synchronize, "
			If trim(strSPerm) <> "" then strSPerm =  left(strSPerm, len(strSPerm)-2)

			If UCase(strdrop) = UCase("/dropInherit") and objAce.AceFlags AND FOLDER_INHERITED_ACE Then
				Wscript.echo "Droped ACL Inheritance " & FolderPerm				
			Else	
				Wscript.echo "Get ACL on Path: " & FolderPerm
				'Date;Time;Folder;Group / User Name;Commom Permission's;Special Permission's;Access Type;Inherit;Error's
				objOutFile.Writeline Date() & ";" & Time() & ";""" & FolderPerm & """;""" & objACE.Trustee.Domain & "\" & objACE.Trustee.Name & """;" & strCPerm & ";" & strSPerm & ";" & strTypePerm  & ";" & strInherit & ";" & strErros
				'wscript.echo objACE.Trustee.Name & " " & objACE.AceFlags
			End if	
		Next
	End If
End Sub

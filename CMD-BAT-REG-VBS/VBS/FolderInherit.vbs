' =============================================================================================================
' List Folders and Child-Folders/SubFolders That Inherit OR Does Not Permission From Parent Folder
' Contact: monimoys@hotmail.com
' =============================================================================================================

Option Explicit

Dim ArrFolders(), IntSize, StrComputer, ObjWMI, ObjWMI2, StartFolderName
Dim ColSubFolders, ColSubFolders2, ObjFolder, ObjFolder2, StrChildFolder
Dim ObjNetwork, StrInput, ObjFSO

Set ObjNetwork = CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName)
Set ObjNetwork = Nothing

WScript.Echo

StartFolderName = vbNullString
WScript.Echo "Enter The Full Path and Folder Name To Start:" & VbCrLf
StartFolderName = WScript.StdIn.Readline

If StartFolderName <> vbNullString AND StrComp(StartFolderName, ".", vbTextCompare) <> 0 Then
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FolderExists(StartFolderName) = False Then
		WScript.Echo "Cannot Continue.Invalid Parameter Entered." & VbCrLf & "You must provide a valid and correct Folder Name."
		Set ObjFSO = Nothing:	WScript.Quit
	End If
	Set ObjFSO = Nothing
	WScript.Echo
	WScript.Echo "Working. Please wait ..."
	WScript.Echo "============================="
	WScript.Echo
	Set ObjWMI = GetObject("WinMgmts:\\" & StrComputer & "\Root\CIMV2")
	IntSize = 0
	
	' -- StartFolderName = "C:\BatTest"
	' -- IntSize = 0

	On Error Resume Next
	Set ColSubFolders = ObjWMI.ExecQuery("Associators of {Win32_Directory.Name='" & StartFolderName & "'} Where AssocClass = Win32_Subdirectory ResultRole = PartComponent")
	ReDim Preserve ArrFolders(IntSize):	ArrFolders(IntSize) = StartFolderName
	IntSize = IntSize + 1
	For Each ObjFolder In ColSubFolders
    		GetSubFolders StartFolderName
	Next
	For Each StrChildFolder In ArrFolders
		CheckNow(StrChildFolder)
	Next
	Set ColSubFolders = Nothing:	Set ObjWMI = Nothing
Else
	WScript.Echo
	WScript.Echo "Cannot Continue." & VbCrLf & "You must provide the Folder Name."
End If
WScript.Quit

' =====================================================================================================================================

Sub GetSubFolders(StrFolderName)
    Set ColSubFolders2 = ObjWMI.ExecQuery("Associators of {Win32_Directory.Name='" & StrFolderName & "'} Where AssocClass = Win32_Subdirectory ResultRole = PartComponent")
    For Each ObjFolder2 In ColSubFolders2
        StrFolderName = Trim(ObjFolder2.Name)
        ReDim Preserve ArrFolders(IntSize):	ArrFolders(IntSize) = StrFolderName
        IntSize = IntSize + 1
        GetSubFolders StrFolderName
    Next
	Set ColSubFolders2 = Nothing
End Sub

Private Sub CheckNow(FolderThis)

	Dim ColFolders, ObjIntFolder, InheritFlag, ObjFLDR
	Dim ObjSD, ObjACE
	
	Set ObjWMI2 = GetObject("WinMgmts:\\" & StrComputer & "\Root\CIMV2")	
	FolderThis = Replace(FolderThis, "\", "\\")
	Set ColFolders = ObjWMI2.ExecQuery("Select * From Win32_Directory Where Name = '" & FolderThis & "'")
	If ColFolders.Count > 0 Then
		For Each ObjIntFolder In ColFolders
			InheritFlag = 0
			Set ObjFLDR = ObjWMI2.Get("Win32_LogicalFileSecuritySetting='" & ObjIntFolder.Name & "'") 
			If ObjFLDR.GetSecurityDescriptor(ObjSD) = 0 Then
				For Each ObjACE In ObjSD.DACL
					If ObjACE.AceFlags AND 16 Then
						InheritFlag = 1 
					End If 
				Next 
			End If
			If InheritFlag = 0 Then
				WScript.Echo "Folder Name: " & ObjIntFolder.Name 
				WScript.Echo "Inherit Permission From Parent: FALSE. Does Not Inherit" 
			Else
				WScript.Echo "Folder Name: " & ObjIntFolder.Name 
				WScript.Echo "Inherit Permission From Parent: TRUE. YES Inherit" 
			End If
			WScript.Echo
			Set ObjSD = Nothing:	Set ObjFLDR = nothing		
		Next
	End If
	Set ColFolders = Nothing:	Set ObjWMI2 = Nothing
	
End Sub

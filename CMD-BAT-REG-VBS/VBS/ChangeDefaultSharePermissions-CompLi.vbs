'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 10/04/08
' Modified : 13/04/08
' Change Default Share Permissions-CompList.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

Dim objFSO, objFile,objTextFile
Dim objDialog,objShell
Dim FileLoc,strComputer,strHexValues,strKeyPath,strValueName,strValue
Dim arrValues,arrHexValues
Dim intResult

'WScript.Shell is used to run an execute
Set objShell = CreateObject("WScript.Shell")
'Scripting.FileSystemObject is used to create and write a text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Opening File
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "Text Files|*.txt|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

Set objTextFile = objFSO.OpenTextFile(FileLoc)

Do Until objTextFile.AtEndOfStream
	strComputer = objTextFile.ReadLine
	ChangeDefaultShare strComputer
Loop

Sub ChangeDefaultShare(strComputer)
'Open the file for writing
Set objFile = objFSO.OpenTextFile("\\Server\EveryoneChangeLogs$\" + strComputer + ".txt",8,True)
'Write starting line
objFile.WriteLine Now & " Checking Default Share Permissions Registry Value."

strHexValues = "01,00,04,80,1c,00,00,00,38,00,00,00,00,00,00,00,14,00,00,00,02,00,08,00,00,00,00,00,01,05,00,00,00,00,00,05,15,00,00,00,8a,5a,41,60,16,c0,ea,32,82,8b,a6,28,1b,5a,0e,00,01,05,00,00,00,00,00,05,15,00,00,00,8a,5a,41,60,16,c0,ea,32,82,8b,a6,28,01,02,00,00"
arrValues = Split(strHexValues,",")

ReDim arrHexValues(UBound(arrValues))
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet\Services\lanmanserver\DefaultSecurity"
strValueName = "SrvsvcDefaultShareInfo"
For i= 0 To UBound(arrValues)
	arrHexValues(i) = Hex_to_Dec(arrValues(i))
Next

oReg.GetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
intErr = 0
If TypeName(strValue) <> "Null" Then
	For i = lBound(strValue) to uBound(strValue)
	    If arrHexValues(i)<>strValue(i) Then _
	    	intErr = intErr + 1
	Next
	If (intErr > 0) And (Ubound(strValue)>0) Then
		objFile.WriteLine "Changing Default Share Permissions Registry Value."
		oReg.SetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrHexValues
	Else
		'objFile.WriteLine "Default Share Permissions Registry Value Exists"
	End If
Else
	objFile.WriteLine  Now & " Changing Default Share Permissions Registry Value."
	oReg.SetBinaryValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrHexValues
End If
objFile.WriteLine Now & " Registry modification completed successfully."
End Sub

Function Hex_to_Dec(hex_value)
	 Hex_to_Dec = CLng("&h" & hex_value)
End Function
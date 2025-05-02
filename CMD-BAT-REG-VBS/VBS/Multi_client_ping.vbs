Option Explicit
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim pFolderName : pFolderName = "clients.txt"

Dim Fso : Set Fso = CreateObject("Scripting.FileSystemObject")
Dim oXLS : Set oXLS = WScript.CreateObject("Excel.Application")
Dim WshShell : Set WshShell = createobject("wscript.shell")

Dim i
Dim intIndex
Dim oFile 
Dim f
Dim png
Dim strComputer()
Dim strIpAddress
Dim strPing
Dim strReply
Dim strRet
Dim pFolder : pFolder = Fso.GetParentFolderName(Wscript.ScriptFullName)

If Not Fso.FileExists(pFolderName) then
	strRet = Msgbox("The file, " & pFolderName & " is not available" & vbCr _
				 & "The file must be located in the same folder as the script." & vbCR _
         & "Please check for the file <<" & pFolderName & ">> in the folder" & vbCR _
         & "<<" & pFolder & ">>")
		Wscript.Quit
End if


 Set oFile = Fso.OpenTextFile(pFolderName, 1)    

''Open and configure Excel
oXLS.Visible = TRUE
oXLS.WorkBooks.Add
oXLS.Columns(1).ColumnWidth = 20
oXLS.Columns(2).ColumnWidth = 20
oXLS.Columns(3).ColumnWidth = 20

''Set column headers
oXLS.Cells(1, 1).Value = "Computer Name"
oXLS.Cells(1, 2).Value = "Return"
oXLS.Cells(1, 3).Value = "IP Address"

''Format text (bold)
oXLS.Range("A1:C1").Select
oXLS.Selection.Font.Bold = True
oXLS.Selection.Interior.ColorIndex = 1
oXLS.Selection.Interior.Pattern = 1 ''xlSolid
oXLS.Selection.Font.ColorIndex = 2
''Left Align text
oXLS.Columns("B:B").Select
oXLS.Selection.HorizontalAlignment = &hFFFFEFDD '' xlLeft


intIndex = 2 ''used in Show sub.
i = 0
Do Until oFile.AtEndOfStream
     Redim Preserve strComputer(i)
     strComputer(i) = oFile.ReadLine
		''Ping Computers
		WshShell.Run("CMD /c ping -n 2 " & strComputer(i) & " >" & pFolder & "\PINGtemp.txt"),0,TRUE
        Set f = fso.OpenTextFile(pFolder & "\PINGtemp.txt", ForReading)
        strPing = f.ReadAll
        f.close

	''NOTE:  The string being looked for in the Instr is case sensitive.  
	''Do not change the case of any character which appears on the 
	''same line as a Case InStr.  AS this will result in a failure.
	Select Case True
		Case InStr(strPing, "Request timed out") > 1 
			strReply = "Request timed out"
			strIpAddress = GetIP(strPing)
		Case InStr(strPing, "could not find host") > 1
			strReply = "Host not reachable"
			strIpAddress = "N/A"
		Case InStr(strPing, "Destination host unreachable") > 1
			strReply = "Host not reachable"
			strIpAddress = GetIP(strPing)		
		Case InStr(strPing, "bytes=") > 1
			strReply = "Ping Successful"
			strIpAddress = GetIP(strPing)
	End Select
	Call Show(strComputer(i), strReply, strIPAddress)
     i = i + 1
Loop

''Delete the PINGtemp file
Fso.DeleteFile pFolder & "\PINGtemp.txt",TRUE

Function GetIP(ByVal reply)
	Dim P
	P = Instr(reply,"[")
	If P=0 Then Exit Function
	reply = Mid(reply,P+1)
	P = Instr(reply,"]")
	If P=0 Then Exit Function
	GetIP = Left(Reply, P-1)
End Function

Sub Show(strName, strValue, strIP)
    oXLS.Cells(intIndex, 1).Value = strName
    oXLS.Cells(intIndex, 2).Value = strValue
    oXLS.Cells(intIndex, 3).Value = strIP
    intIndex = intIndex + 1
    oXLS.Cells(intIndex, 1).Select
End Sub

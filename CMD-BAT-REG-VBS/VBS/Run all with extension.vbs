Dim myFolder : myFolder = "c:\temp"
 
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
 
For Each file In fso.GetFolder(myFolder).Files
    If Len(file.Name) > 5 Then
        Dim extension : extension = UCase(Right(file.Name, 3))
        Select Case extension
        Case "VBS":
            sh.Run "wscript """ & file.Path & """", 1, True
        Case "EXE":
        Case "BAT":
            sh.Run file.Path, 1, True
        End Select
    End If
Next
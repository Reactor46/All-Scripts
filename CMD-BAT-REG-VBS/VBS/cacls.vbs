
Dim oShell, FoldPerm, Calcds, oFSO

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

sSysDir = oFSO.GetSpecialFolder(1).Path
If Right(sSysDir,1) <> "\" Then sSysDir = sSysDir & "\"

Calcds = sSysDir & "cacls.exe" 

'Chang The folder Name, User and Access rights in the following line of code  

FoldPerm = """" & Calcds &"""" & """C:\oracle\BIToolsHome_1\discoverer""" & " /E /T /C /G " & """Authenticated Users""" & ":C" 

oShell.Run FoldPerm, 1 ,True
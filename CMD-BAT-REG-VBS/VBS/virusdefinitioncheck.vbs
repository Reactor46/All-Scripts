Option Explicit
Dim VirusDefCfg, FileSys, FSO, LastModified, DateDifference, noSymantecPresent 
VirusDefCfg = "C:\ProgramData\Symantec\Definitions\VirusDefs\definfo.dat"
noSymantecPresent = 9999
Set FileSys = CreateObject("Scripting.FileSystemObject")
Set FSO = CreateObject("Scripting.FileSystemObject")
If FileSys.FileExists(VirusDefCfg) <> True Then
WScript.Echo noSymantecPresent
WScript.Quit
End If

LastModified = FSO.GetFile(VirusDefCfg).DateLastModified
DateDifference = DateDiff("d", LastModified, Now())
WScript.Echo DateDifference

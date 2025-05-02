'Install Script for TSTool3 - Matthew Painter 18-Sept-2012




'1. Ensure application 'TSTool3.ps1' is placed on a central server share along with Icon File 'TSTool3.ico' and this install script 'install.vbs'. 
'2. If you have an SCCM server, edit 'GlobalSettings.csv' with your SCCM server name and name space and place this file in the same server share.
'3. Distribute Full UNC path to 'install.vbs' rather than directly to TSTool3.ps1' as app works best from the shortcut 'TSTool3.LNK' that this script creates.




Set objShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(Wscript.ScriptFullName)
ServerShare = objFSO.GetParentFolderName(objFile) 




DTsLinkFile = objShell.SpecialFolders("Desktop")&"\TSTool3.LNK"
SMsLinkFile = objShell.SpecialFolders("StartMenu")&"\Programs\TSTool3.LNK"
TargetP = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
icon = ServerShare&"\TSTool3.ico"
args = "-executionpolicy bypass -sta -file "&chr(34)&ServerShare&"\TSTool3.ps1"&chr(34)




If objFSO.FileExists(DTsLinkFile) Then
    IQuestion = MsgBox ("Desktop Icon Already Exists"&vbcrlf&"Replace?", 52, "Question") 
Else
    IQuestion = MsgBox ("Install TSTool3 ShortCut to Desktop", 52, "Question")    
End If

if IQuestion = 6 then 
   MyShortCut DTsLinkFile, TargetP, icon, args
end if




If objFSO.FileExists(SMsLinkFile) Then
    IQuestion = MsgBox ("Start Menu Icon Already Exists"&vbcrlf&"Replace?", 52, "Question")
Else
    IQuestion = MsgBox ("Install TSTool3 ShortCut to StartMenu", 52, "Question") 
End If

if IQuestion = 6 then 
    MyShortCut SMsLinkFile, TargetP, icon, args
end if




Sub MyShortCut(LinkFile, TargetP, icon, args)
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set oLink = objShell.CreateShortcut(LinkFile)
    oLink.TargetPath = TargetP
    oLink.Arguments = args
    oLink.Description = "W7 Desktop Support Tool"
    oLink.WorkingDirectory = "C:\Windows\System32\WindowsPowerShell\v1.0\"
    oLink.IconLocation = icon
    oLink.Save
End sub

'JRE_Updater 
'Updates and optionally deploys or removes Java Runtime 
'Last Update: 06/12/2015 
'By Kevin Denham (kevin.denham@gmail.com) 
' 
'=================   Options   ================================== 
'Download=0 if this is a client that is 'not' downloading 
Download=1 
'ModifyJRE=0 if this is a host that is 'only' downloading 
ModifyJRE=1 
 
'Location for downloaded files.  Example: FileStore="\\Server\SharedFolder\" 
'Leaving FileStore as "" uses the script's working directory 
FileStore=""  
 
'0 = Remove JRE, 1 = Update Existing JRE, 2 = Deploy and/or Update JRE 
x86=1 
x64=1 'For x64 bit browsers otherwise use only x86 for x64 systems 
 
'Add any extra Java CLI switches 
'http://docs.oracle.com/javase/8/docs/technotes/guides/install/windows_installer_options.html#A1097528 
'URL Updated for Java 8 Options 
'Example: switches="/s INSTALLDIR=D:\Program Files(x86)\Java\" 
switches32="/s" 
switches64="/s" 
 
On error resume next
 

'Do not edit below this line 
'=================================================================== 
Set App = CreateObject("Shell.Application") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set oShell = WScript.CreateObject("WScript.Shell")  
Set installer = CreateObject("WindowsInstaller.Installer") 
Set req = CreateObject("MSXML2.XMLHTTP.6.0") 
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0") 
  
areAdmin = false 
 
'If FileStore was not specified use the script's working directory 
If FileStore = "" Then FileStore = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\" 
file32 = FileStore & "JRE_32bitBrowsers.exe" 
file64 = FileStore & "JRE_64bitBrowsers.exe" 
 
 
If Download = 1 Then 
    If Not objFSO.FolderExists(FileStore) Then objFSO.CreateFolder(FileStore) 
    logFile = FileStore & "JRE_Bundle.log" 
 
    'Create the log that records the last download bundle ID 
    If Not objFSO.FileExists(logfile) Then 
        header = "Current Java Bundle Log: "  
        objFSO.OpenTextFile(logFile, 8, True).WriteLine header  
    End if 
 
    Set openLog = objFSO.OpenTextFile(logFile, 1, True) 
    readLog = openLog.ReadAll 
 
    'Query java.com 
    req.open "GET", "http://www.java.com/en/download/manual.jsp", False 
    req.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64)" 
    req.send()  
    'Parse query response 
    Call parseCheck(" Offline", file32, bundle32) 
    Call parseCheck(" (64-bit)", file64, bundle64) 
End If 
 
Function parseCheck(str, str2, bndl) 
    strArr = split(req.responseText, "Download Java software for Windows" & str) 
    strArr2 = split(strArr(1), Chr(34))  
    bndl = strArr2(2) 
    If instr(readLog, bndl) = false Then 
        objFSO.OpenTextFile(logFile, 8, True).WriteLine Date & "|" & bndl 
        'Our bundle is out-of-date, download latest 
        Call getBundle(str2, bndl) 
    End If  
End Function 
 
Function getBundle(fName, bURL) 
    objXMLHTTP.open "GET", bURL, false  
    objXMLHTTP.send()  
    If objXMLHTTP.Status = 200 Then  
        Set objADOStream = CreateObject("ADODB.Stream")  
        objADOStream.Open  
        objADOStream.Type = 1  
        objADOStream.Write objXMLHTTP.ResponseBody  
        objADOStream.Position = 0    
 
        If objFSO.FileExists(fName) Then objFSO.DeleteFile fName  
        objADOStream.SaveToFile fName  
        objADOStream.Close  
    End if  
End Function 
 
'Compare version numbers 
Function chkArr(arr1, arr2, i) 
    If (CLng(arr1(i)) < (CLng(arr2(i)))) Then  
        chkArr = True 
        Exit Function 
    End If     
    If (UBound(arr1) > i) And (UBound(arr2) > i) Then chkArr = chkArr(arr1, arr2, i + 1) 
End Function 
 
If ModifyJRE = 1 Then 
    If objFSO.FileExists(file32) Then store32 = objFSO.GetFileVersion(file32)  
    If objFSO.FileExists(file64) Then store64 = objFSO.GetFileVersion(file64)  
     
    'Search for old Java installs 
        For Each product In installer.Products 
        'If product is JRE 
        If InStr(product, "26A24AE4-039D-4CA4-87B4-") Then 
            Version = installer.ProductInfo(product, "VersionString") 
            'If installed JRE is 32 bit 
            If Mid(product,29,2) = "32" Then 
                'If installed JRE is out-of-date or we specified removal 
                If (chkArr(Split(Version, "."), Split(store32, "."), CInt(0)) = True) Or x86 = 0 Then 
                    If areAdmin = false Then Call IsElevated 
                    oShell.run("msiexec.exe /x " & product & " /qn"), 1, true 
                    If x86 = 1 Then x86 = 2     
                Else If x86 = 2 Then x86 = 1                      
                End If  
            End If 
            If Mid(product,29,2) = "64" Then 
                If (chkArr(Split(Version, "."), Split(store64, "."), CInt(0)) = True) Or x64 = 0 Then 
                    If areAdmin = false Then Call IsElevated 
                    oShell.run("msiexec.exe /x " & product & " /qn"), 1, true 
                    If x64 = 1 Then x64 = 2 
                Else If x64 = 2 Then x64 = 1 
                End If 
            End If 
        End if  
    Next 
    'If FileStore is UNC, transfer files 
    If  (Mid(FileStore,1,2) = "\\") And ((x86 = 2) Or (x64 = 2)) Then 
        temp = oShell.ExpandEnvironmentStrings("%temp%") 
        Set TypeLib = CreateObject("Scriptlet.TypeLib") 
        tmpFolder = objFSO.CreateFolder(temp & "\" & TypeLib.Guid) 
        objFSO.CopyFile FileStore & "*", tmpFolder, True 
        file32 = tmpFolder & "\" & objFSO.GetFileName(file32) 
        file64 = tmpFolder & "\" & objFSO.GetFileName(file64) 
    End If 
    'Install the latest JRE 
    If x86 = 2 And objFSO.FileExists(file32) Then 
        If areAdmin = false Then Call IsElevated 
        oShell.run(chr(34) & file32 & chr(34) & " " & switches32), 1, true 
    End If 
    If x64 = 2 And objFSO.FileExists(file64) Then  
        If areAdmin = false Then Call IsElevated 
        oShell.run(chr(34) & file64 & chr(34) & " " & switches64), 1, true 
    End If 
    If objFSO.FolderExists(tmpFolder) Then objFSO.DeleteFolder tmpFolder, True 
    'Uninstall the Java Update Scheduler 
    For Each product In installer.Products 
        If InStr(product, "4A03706F-666A-4037-7777-5F2748764D10") Then 
            If areAdmin = false Then Call IsElevated 
            oShell.run("msiexec.exe /X{4A03706F-666A-4037-7777-5F2748764D10} /qn"), 1, true 
        End if 
    Next 
End If 
 
Function IsElevated 
    IsElevated = CreateObject("WScript.Shell").Run("cmd.exe /c ""whoami /groups|findstr S-1-16-12288""", 0, true) = 0 
    If IsElevated = false Then 
        App.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas" 
        wscript.quit  
    Else 
        areAdmin = true 
    End If 
End function 
'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

Option Explicit

Dim ObjFSO : Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Dim ObjShell : Set ObjShell = CreateObject("WScript.Shell")
Dim WshNetwork : Set WshNetwork = CreateObject("WScript.Network")

' Run cmd certutil -urlcache * delete 
Sub DeleteUrlCache(strComputer)
    On Error Resume Next
    
    Dim objWMIService, colSettings
    
    ObjShell.run "certutil -urlcache * delete " ,vbhide
    
    ' Check process
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_Process where Name='certutil.exe'")
    While colSettings.Count <> 0   
        Set colSettings = objWMIService.ExecQuery _
        ("Select * from Win32_Process where Name='certutil.exe'")
    Wend
End Sub 

' Get the Windows directory of the remote computer
Function GetRemoteWindowsDirectories(strComputer)
    On Error Resume Next
    
    Dim objWMIService, colItems, userArrayList, objItem
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
    
    Set userArrayList = CreateObject( "System.Collections.ArrayList" )
    
    For Each objItem In colItems
        userArrayList.Add objItem.WindowsDirectory	
    Next
    
    Set GetRemoteWindowsDirectories = userArrayList
End Function

' Delete files from the specific path
Sub DeleteFolderContents(strFolderPath)
    On Error Resume Next
    
    Dim folder, file, subFolder
    
    If ObjFSO.FolderExists(strFolderPath) Then 
        Set folder = ObjFSO.GetFolder(strFolderPath)
        
        ' Delete files in the specific folder
        For Each file In folder.Files
            file.Delete
        Next
        
        ' Delete subfolders in the specific folder
        For Each subFolder In folder.SubFolders
            subFolder.Delete
        Next
        
        WScript.Echo "The path: " & strFolderPath & " has been cleared!"
    Else 
        WScript.Echo "The path: " & strFolderPath & " don't exist!"
    End If 
End Sub

' Delete files from the specific computer and path
Sub RemoteDeleteFolderContents(strComputer, strWindowsDirectory, strFolderPath)
    On Error Resume Next
    Err.Clear
    
    Dim mapDrive, folderPath  
    mapDrive = Left(strWindowsDirectory,1)
    folderPath = "Y" & Right(strWindowsDirectory,Len(strWindowsDirectory)-1) & strFolderPath
    
    ' Mapping remote computer disk drive
    WshNetwork.MapNetworkDrive "Y:", "\\" & strComputer & "\" & mapDrive & "$"
    
    If Err.Number <> 0 Then    
        WScript.Echo "Mapping Drive failed!"
    Else
        Call DeleteFolderContents(folderPath)
    End If
    
    WshNetwork.RemoveNetworkDrive "Y:",True,True    
End Sub

' Fix issues on the computer
Sub FixIssues(strComputer)
    On Error Resume Next
    
    Call DeleteUrlCache(strComputer)
    
    Dim winDirs : Set winDirs = GetRemoteWindowsDirectories(strComputer)
    Dim winDir	
    For Each winDir In winDirs
        'LocalService:
        Call RemoteDeleteFolderContents(strComputer, winDir, "\ServiceProfiles\LocalService\AppData\LocalLow\Microsoft\CryptnetUrlCache\Content")
        Call RemoteDeleteFolderContents(strComputer, winDir, "\ServiceProfiles\LocalService\AppData\LocalLow\Microsoft\CryptnetUrlCache\MetaData")
        
        'NetworkService: 
        Call RemoteDeleteFolderContents(strComputer, winDir, "\ServiceProfiles\NetworkService\AppData\LocalLow\Microsoft\CryptnetUrlCache\Content")
        Call RemoteDeleteFolderContents(strComputer, winDir, "\ServiceProfiles\NetworkService\AppData\LocalLow\Microsoft\CryptnetUrlCache\MetaData")
        
        'LocalSystem: 
        Call RemoteDeleteFolderContents(strComputer, winDir, "\System32\config\systemprofile\AppData\LocalLow\Microsoft\CryptnetUrlCache\Content")
        Call RemoteDeleteFolderContents(strComputer, winDir, "\System32\config\systemprofile\AppData\LocalLow\Microsoft\CryptnetUrlCache\MetaData")
    Next
End Sub

' Bulk remote fix the issues for the computers.
Sub BulkRemoteFixIssues(strPCNamePath)
    On Error Resume Next
    
    Dim pcNameFile, pcName
    
    If Not ObjFSO.FileExists(strPCNamePath) Then 
        WScript.Echo "The path of the file is invalid!"
        Exit Sub
    End If 
    Set pcNameFile = ObjFSO.OpenTextFile(strPCNamePath,1)
    Do Until pcNameFile.AtEndOfStream
        pcName = pcNameFile.ReadLine
        WScript.Echo "Now dealing with the computer: " & pcName
        Call FixIssues(pcName)
    Loop
    pcNameFile.Close 
End Sub
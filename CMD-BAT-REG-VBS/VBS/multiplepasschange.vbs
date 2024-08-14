On Error Resume Next 
Const ForAppending = 8 
Const ForReading = 1 
strNewAccountName="admin!!!"
strNewPassword = "Welcome123"

'Declaring the variables 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set SrvList = objFSO.OpenTextFile("Server_List.txt", ForReading)

Do Until SrvList.AtEndOfStream 
    StrComputer = SrvList.ReadLine
    
set objComp = GetObject("WinNT://" & strComputer)
set objUser = GetObject("WinNT://" & strComputer & "/pcadmin,user")
objuser.Setpassword strNewPassword
objUser.SetInfo
set objNewUser = objComp.MoveHere(objUser.ADsPath,strNewAccountName)
Loop 





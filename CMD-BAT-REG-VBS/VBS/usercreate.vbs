On Error Resume Next 
Const ForAppending = 8 
Const ForReading = 1 
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const ADS_PROPERTY_UPDATE = 2
strNewAccountName="%Admin!!!"
strNewPassword = "Welcome123"
strUser = "%Admin!!!"

'Declaring the variables 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set SrvList = objFSO.OpenTextFile("Server_List.txt", ForReading)

Do Until SrvList.AtEndOfStream 
    StrComputer = SrvList.ReadLine
    
set objComp = GetObject("WinNT://" & strComputer)
Set colAccounts = GetObject("WinNT://" & strComputer)
Set objUser = colAccounts.Create("user" , "%Admin!!!")
objuser.Setpassword strNewPassword
set objNewUser = objComp.MoveHere(objUser.ADsPath,strNewAccountName)

objUserFlags = objUser.Get("UserFlags")
objPasswordExpirationFlag = objUserFlags OR ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag 
objUser.SetInfo

Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser & ",user")
objGroup.Add(objUser.ADsPath)
Loop 




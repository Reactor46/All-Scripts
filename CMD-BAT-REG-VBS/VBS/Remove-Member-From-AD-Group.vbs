Const ADS_PROPERTY_DELETE = 4  
 
Set oFSO = CreateObject("Scripting.FileSystemObject") 

sFile = "c:\temp\users.txt" 
 If oFSO.FileExists(sFile) Then 
  Set oFile = oFSO.OpenTextFile(sFile, 1) 
   Do While Not oFile.AtEndOfStream 
    strUsername = oFile.ReadLine 

Set objGroup = GetObject _ 
   ("LDAP://cn=gg-test,OU=Application-Role-Profiles,OU=Groups,dc=testing,dc=com")  
  
objGroup.PutEx ADS_PROPERTY_DELETE, _ 
    "member",Array("cn=" & strUsername & ",OU=Group,OU=Accounts,DC=testing,DC=com") 
objGroup.SetInfo
Loop
End If
mailDomain = ""
firstRun = True


Dim rootDSE, domainObject, adDomain, mailDomain
Set rootDSE = GetObject("LDAP://RootDSE")
domainContainer = rootDSE.Get("defaultNamingContext")
Set domainObject = GetObject("LDAP://" & domainContainer)


Set fs = CreateObject ("Scripting.FileSystemObject")
Set ousFile = fs.CreateTextFile (".\OUs.txt")
Set usersFile = fs.CreateTextFile (".\Users.csv")
Set groupsFile = fs.CreateTextFile (".\Groups.csv")

MsgBox "Running, this will take a couple of minutes..." & vbCRLF & vbCRLF & "Please hit OK to continue"
exportUsers(domainObject)


Set oDomain = Nothing
MsgBox "Finished"
WScript.Quit



Sub ExportUsers(oObject)
   Dim oAD
   For Each oAD in oObject
      Select Case oAD.Class
         Case "user"
            usersFile.WriteLine oAD.distinguishedName
            usersFile.WriteLine chr(34) & oAD.givenName & chr(34) & "," & chr(34) & oAD.sn & chr(34) & "," & chr(34) & oAD.displayName & chr(34) & "," & chr(34) & oAD.userPrincipalName & chr(34) & "," & chr(34) & oAD.sAMAccountName & chr(34)
         Case "organizationalUnit"
            ousFile.WriteLine oAD.distinguishedName
            ExportUsers(oAD)
         Case "container"
            ExportUsers(oAD)
         Case "group"
            groupsFile.WriteLine "group," & oAD.sAMAccountName & "," & oAD.groupType & "," & oAD.distinguishedName 

            oAD.GetInfo
            On Error Resume Next
            arrMemberOf = oAD.GetEx("member")
            If Err.Number = 0 then
               For Each strMember in arrMemberOf
                  groupsFile.WriteLine "member," & strMember
               Next           
            Else
               Err.Clear
            End If
            On Error Goto 0
      End Select
   Next
End Sub


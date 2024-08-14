' Author: Romano Jerez, http://www.netvnext.com
' November 2012


strADgroup = "LDAP://CN=ADGroup-LocalAdmins,OU=SecurityGroups,DC=NETvNext,DC=COM"
Set objADGroup = GetObject(strADgroup)

strComputer = "."
Set objLocalGroup = GetObject("WinNT://" & strComputer & "/Administrators")


For Each objLocalMember In objLocalGroup.Members

   If objLocalMember.Class = "User" Then
      
       bAuthorized = 0

       For Each objADMember in objADGroup.Members
          If objADMember.Class = "user" Then
             If objADMember.samAccountName = objLocalMember.Name Then
                bAuthorized = 1
             End If
          End If
       Next

       If objLocalMember.Name = "Administrator" Then
          bAuthorized = 1
       End If

       If objLocalMember.Name = "Admin2" Then
          bAuthorized = 1
       End If

       If bAuthorized = 0 Then
          WScript.Echo "Removing " & objLocalMember.Name
          objLocalGroup.Remove(objLocalMember.AdsPath)
       End If
   End If

Next


Set objADGroup = Nothing
Set objLocalGroup = Nothing
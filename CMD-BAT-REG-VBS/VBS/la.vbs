' ------------------------------------------------------------------------------------------------------------------------------------
' - Script:  la.vbs
' - Date:    08/26/2013
' - Author:  Jeff Mason / aka bitdoctor / aka tnjman
' - Purpose: To compare and report on AUTHORIZED and UNAUTHORIZED local admins on remote computers
' - Assumptions: 
' -   1) The UserID under which you run this has admin rights on the Remote computer
' -
' - How to run: 
' -   1) Save this as c:\scripts\la.vbs
' - 
' -   2) Create "la.bat" file with names of remote computers for which you want to audit the "local admins" group, like this
' -      REM - c:\scripts\la.bat - put each computer (workstation/server) name on a separate line
' -      cscript /nologo c:\scripts\la.vbs remote-pc1
' -      cscript /nologo c:\scripts\la.vbs remote-pc2
' -      cscript /nologo c:\scripts\la.vbs remote-server1
' -
' -   3) Execute the "la.bat" file (interactively or via schedule), redirecting output to LOG/report file:
' -      c:\scripts\la.bat > c:\scripts\la-log.txt 2>&1
' -
' -   4) Open the resuling log file via Excel as "text," delimited by "comma."
' -      You now have a list of all members of the local "administrators" group from the list of remote PCs in the "la.bat" file
' -      You can now sort the Excel file by "Good," "Bad," etc.; so you can remediate computers containing unauthorized admins.
' -
' ------------------------------------------------------------------------------------------------------------------------------------
'
on error resume next
strComputer = Wscript.Arguments.Item(0)
If strComputer = "" Then
 strComputer = "."
End If
call CompareLocalAdmins(strComputer)
WScript.Quit 0

'--------------------------------------------------
'- Compare Local Admins to List of Valid Accounts -
'--------------------------------------------------

Sub CompareLocalAdmins(RemoteSystem)
Dim objComp
strComputer = RemoteSystem
Set objComp = GetObject("WinNT://" & strComputer)
objComp.GetInfo
If objComp.PropertyCount > 0 Then
  Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
  If objGroup.PropertyCount > 0 Then
'    WScript.Echo "QUESTIONABLE members of local Admins group on " & strComputer & " are:"
    okflag = "Y"
    For Each mem In objGroup.Members
      memx = LCase(mem.Name)
      'List of locally-authorized admins against which to compare
      If (memx = "workstation admins" or memx = "domain admins" or memx = "special-admin" or memx = "administrator") Then
        'WScript.Echo "A-OKAY!"
        'Note: "special-admin" above is an example of a locally-authorized admin
      Else
        okflag = "N"
        WScript.Echo strComputer & "," & mem.Name & ",BAD" 
      End If
    Next
      If okflag = "Y" Then
        WScript.Echo strComputer & ",ALLOK"
      End If
  Else
    WScript.echo "Unable to check remote computer: " & strComputer
    WScript.Quit 1
  End If
Else
  WScript.Echo "Unable to check remote computer: " & strComputer
  WScript.Quit 1
End If

End Sub
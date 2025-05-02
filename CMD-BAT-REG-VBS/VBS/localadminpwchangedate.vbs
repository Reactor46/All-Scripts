'LOCAL ADMINISTRATOR ACCOUNT LAST PASSWORD CHANGE DATE
'By Rob Cilia 7/12/2013

WScript.Echo "Local Administration Account Password Change Date Report"
WScript.Echo "Current Date & Time: " & Now()
WScript.Echo " "

On Error Resume Next

Set colComputers = GetObject("LDAP://OU=Servers,DC=domain,DC=com")
For Each objComputer in colComputers
      strComputer = objComputer.CN
      Set objUser = GetObject("WinNT://" & strComputer & "/administrator")
      intPasswordAge = objUser.PasswordAge
      intPasswordAge = intPasswordAge * -1 
      dtmChangeDate = DateAdd("s", intPasswordAge, Now)

If Err.Number = 0 Then
    WScript.Echo objComputer.CN & "\ranger20b   " & "Password last changed: " & dtmChangeDate
    WScript.Echo " "
Else
    WScript.Echo "Error communicating with: " & objComputer.CN
    WScript.Echo "Machine is not reachable or computer account is disabled, data is not available for this machine - If the computer account is disabled move to the Disabled\Computers OU."
    WScript.Echo " "
    Err.Clear
End If
      
Next
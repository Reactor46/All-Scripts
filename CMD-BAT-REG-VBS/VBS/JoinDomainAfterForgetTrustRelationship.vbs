
On Error Resume Next

Const JOIN_DOMAIN             = 1
Const ACCT_CREATE             = 2
Const ACCT_DELETE             = 4
Const WIN9X_UPGRADE           = 16
Const DOMAIN_JOIN_IF_JOINED   = 32
Const JOIN_UNSECURE           = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET        = 256
Const INSTALL_INVOCATION      = 262144

strDomain   = "CONTOSO"
strPassword = "J01ndreD0maine"
strUser     = "!joindredomaine"

Set objNetwork = CreateObject("WScript.Network")

strComputer = objNetwork.ComputerName

Set objComputer = _
    GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
    strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" _
    & strComputer & "'")

ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, _
    strPassword, _
    strDomain & "\" & strUser, _
    NULL, _
    JOIN_DOMAIN + ACCT_DELETE + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED)

if Err <> 0 then
    WScript.Echo "Join failed with error: " & ReturnValue
else
    WScript.Echo "Successfully joined " & strComputer & " to " & strDomain
end if

    WScript.Echo date & "   " & time & "  : ReturnValue = " & ReturnValue

If ReturnValue = 5 then
    WScript.Echo "Access was Denied for adding the Computer to the Domain"
End If
If ReturnValue = 87 then
    WScript.Echo "The parameter is incorrect"
End If
If ReturnValue = 110 then
    WScript.Echo "The system cannot open the specified object"
End If
If ReturnValue = 1323 then
    WScript.Echo "Unable to update the password"
End If
If ReturnValue = 1326 then
    WScript.Echo "Logon failure: unknown username or bad password"
End If
If ReturnValue = 1355 then
    WScript.Echo "The specified domain either does not exist or could not be contacted"
End If
If ReturnValue = 2224 then
    WScript.Echo "Account exists in the Domain"
End If
If ReturnValue = 2691 then
    WScript.Echo "The machine is already joined to the domain"
End If
If ReturnValue = 2692 then
    WScript.Echo "The machine is not currently joined to a domain"
End If
if ReturnValue > 2692 then
    WScript.Echo "The Computer was sucssefully added to the domain"
Else
	if ReturnValue = 0 then
    		WScript.Echo "The Computer was sucssefully added to the domain"
	Else
    		WScript.Echo "The Computer was NOT added to the domain"
	End if
End if

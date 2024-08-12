set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchExchangeServer);name,serialNumber,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
Wscript.echo "Exchange Servers Versions IMF Updates"
Wscript.echo
While Not Rs.EOF
arrSerial = rs.Fields("serialNumber")
For Each Serial In arrSerial
strexserial = Serial
Next
call getIMFversion(rs.fields("name"))

Rs.MoveNext
Wend
Rs.Close
Conn.Close
Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing

function getIMFversion(strComputer)
on error resume next
const HKEY_LOCAL_MACHINE = &H80000002
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Updates\Exchange Server 2003\SP3\KB907747"
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "PackageVersion", Imfver
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, "InstalledDate", Imfinstdate
if isnull(Imfver) then
	wscript.echo strComputer  & " : No IMF Updates Installed"
	wscript.echo 
else
	wscript.echo strComputer  & " : " & Imfver
	wscript.echo "Update Installed on : " & Imfinstdate
	wscript.echo
end if

end function

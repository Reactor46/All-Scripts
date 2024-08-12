set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchExchangeServer);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
Wscript.echo "Exchange Servers Versions"
Wscript.echo
While Not Rs.EOF
	ExVersion = getExversion(rs.fields("name"))
	wscript.echo rs.fields("name") & " " & ExVersion
	Rs.MoveNext
Wend
Rs.Close
Conn.Close
Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing

function getExversion(strComputer)
const HKEY_LOCAL_MACHINE = &H80000002
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
    & strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
	select case ucase(subkey)
		case ucase("DB20F7FD-67BC-4813-8808-78F63E89EB56") ExVersion = "Exchange 2000 Standard"
		case ucase("775CF3DA-C007-4709-B4CC-CE2239BE2E03") Exversion = "Exchange 2000 Standard"
		case ucase("FC6FA539-452D-4a9b-8065-C1FA74B86F83") ExVersion = "Exchange 2000 Standard Evaluation"
		case ucase("D3574E0C-360A-44d5-858C-33323C2D79F2") ExVersion = "Exchange 2000 Enterprise" 
		case ucase("F8567801-906B-439b-8D6A-87BDFEC9BA52") ExVersion = "Exchange 2000 Enterprise" 
		case ucase("65D9643D-06E8-47d6-865E-80F4CC9BB879") ExVersion = "Exchange 2000 Enterprise" 
		case ucase("8B102332-6052-4af3-ADFA-35A3DED0506A") ExVersion = "Exchange 2000 Enterprise Evaluation" 
		case ucase("ee2d3727-33c0-11d2-ab50-00c04fb1799f") ExVersion = "Exchange 2000 Standard SBS" 
                case ucase("EAE76D62-2691-4883-8BBB-1F2EE6D370D5") ExVersion = "Exchange 2003 Standard"	
		case ucase("9682A75B-EBD1-4c7d-88F9-13BE236F706C") ExVersion = "Exchange 2003 Standard" 
		case ucase("9161A261-6ABE-4668-BBFA-AD06B3F642CF") ExVersion = "Exchange 2003 Standard" 
		case ucase("D8862944-4F8A-429d-9A4F-6F201428FB0C") ExVersion = "Exchange 2003 Standard Evaluation" 
		case ucase("C160866F-DE53-434f-ADF1-CC42ABBF8778") ExVersion = "Exchange 2003 Standard Evaluation"
		case ucase("74F3BB3C-A434-48fa-AAC1-3FC37CD2B0DB") ExVersion = "Exchange 2003 Enterprise"
		case ucase("7F4210A8-D3B4-480a-835E-D5DAA0D0B157") ExVersion = "Exchange 2003 Enterprise" 
		case ucase("4050D45F-9598-44bc-8C07-4C1BBE22EFBB") ExVersion = "Exchange 2003 Enterprise"
		case ucase("F95DE19F-CF69-4b03-81B6-9EC050D20D3B") ExVersion = "Exchange 2003 Enterprise" 
		case ucase("3D5A0E1C-B6DA-42a7-A871-03CD2E30FEA3") ExVersion = "Exchange 2003 Enterprise Evaluation"
		case ucase("2B8EC4BD-5638-47e2-8817-1A50B38A828D") ExVersion = "Exchange 2003 Enterprise Evaluation" 
		case ucase("5717D53E-DD6D-4d1e-8A1F-C7BE620F65AA") ExVersion = "Exchange 2003 Standard SBS" 
	end select
Next
getExversion = ExVersion
end function







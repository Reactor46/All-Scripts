if WScript.Arguments.Count <> 2 then
  	call DisplayUsage
else
	call Main()
end if


sub main()
servername = wscript.arguments(0)
showarg = lcase(wscript.arguments(1))
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS ADDisplayName, " & _
           "  NEW adVarChar(255) AS ADLegacyDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS WMIDisplayName, " & _
           "      NEW adVarChar(255) AS WMILegacyDN, " & _
           "      NEW adVarChar(255) AS WMILoggedOnUserAccount, " & _
           "      NEW adVarChar(255) AS WMIClientVersion, " & _
           "      NEW adVarChar(255) AS WMIClientIP, " & _
           "      NEW adVarChar(255) AS WMIClientMode " & _
	   ")" & _
           "      RELATE ADLegacyDN TO WMILegacyDN) AS rsADWMI " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof		
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs2 = Com.Execute
	while not rs2.eof
		objParentRS.addnew 
		objParentRS("ADDisplayName") = rs2.fields("displayname")
		objParentRS("ADLegacyDN") = rs2.fields("legacyExchangeDN")
		objParentRS.update	
		rs2.movenext
	wend
	wscript.echo "finished 1st AD of Mailbox's"
	rs.movenext
wend

strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername &"/root/MicrosoftExchangeV2"
Select case showarg
	case "owa" sqlstate = "Select * FROM Exchange_Logon where StoreType = 1 and ClientVersion = 'HTTP'"
	case "outlook" sqlstate = "Select * FROM Exchange_Logon where StoreType = 1 and ClientVersion <> 'SMTP' AND ClientVersion <> 'OLEDB' AND ClientVersion <> 'HTTP'"
	case else sqlstate = "Select * FROM Exchange_Logon where StoreType = 1 and ClientVersion <> 'SMTP' AND ClientVersion <> 'OLEDB'"
end select
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_ExchangeLogons = objWMIExchange.ExecQuery(sqlstate,,48)
objChildRS.LockType = 3
Set objChildRS = objParentRS("rsADWMI").Value
For each objExchange_ExchangeLogon in listExchange_ExchangeLogons
	if objExchange_ExchangeLogon.LoggedOnUserAccount <> "NT AUTHORITY\SYSTEM" then
			objChildRS.addnew 
			objChildRS("WMIDisplayName") = objExchange_ExchangeLogon.MailboxDisplayName
			objChildRS("WMILegacyDN") = objExchange_ExchangeLogon.MailboxLegacyDN
			objChildRS("WMILoggedOnUserAccount") = objExchange_ExchangeLogon.LoggedOnUserAccount
			objChildRS("WMIClientVersion") = objExchange_ExchangeLogon.ClientVersion
			objChildRS("WMIClientIP") = objExchange_ExchangeLogon.ClientIP
			objChildRS("WMIClientMode") = objExchange_ExchangeLogon.ClientMode
			objChildRS.update
	end if
Next
wscript.echo "finished Exchange WMI query"
Wscript.echo "Dislay Results " & showarg
wscript.echo
objParentRS.MoveFirst
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("rsADWMI").Value 
    objChildRS.sort = "WMILoggedOnUserAccount"
    select case showarg
	case "all"	Wscript.echo objParentRS.fields("ADDisplayName")
    			if objChildRS.recordcount = 0 then Wscript.echo "		" & "Not Logged On"
	case "loggedon" if objChildRS.recordcount <> 0 then Wscript.echo objParentRS.fields("ADDisplayName")
	case "owa" if objChildRS.recordcount <> 0 then Wscript.echo objParentRS.fields("ADDisplayName")
	case "outlook" if objChildRS.recordcount <> 0 then Wscript.echo objParentRS.fields("ADDisplayName")
	case "loggedout"if objChildRS.recordcount = 0 then Wscript.echo  objParentRS.fields("ADDisplayName")
	case else	Wscript.echo objParentRS.fields("ADDisplayName")
    end select	
    if showarg <> "loggedout" then
    Do While Not objChildRS.EOF
	currec = objChildRS.fields("WMILoggedOnUserAccount") & objChildRS.fields("WMIClientVersion") & objChildRS.fields("WMIClientIP") & objChildRS.fields("WMIClientMode")
	if currec <> prevrec then 
		Wscript.echo "		" & objChildRS.fields("WMILoggedOnUserAccount") & "   " & getversion(objChildRS.fields("WMIClientVersion")) & "  " &  _
		getmode(objChildRS.fields("WMIClientMode")) & "  " & objChildRS("WMIClientIP") 
	end if
	prevrec = currec
	objChildRS.MoveNext
    Loop
    end if   		
    objParentRS.MoveNext
Loop

end sub


function getversion(version)

Select case version
        case "4.0.994.0"  getversion = "Outlook 97 Initial release in Q1/1997."
        case "5.0.1457.0"  getversion = "Outlook 97 Ships only with Microsoft Exchange 5.0 Service Pack 1"
        case "5.0.1458.0"  getversion = "Outlook 97 Ships only with Microsoft Office 97 SR-1"
        case "5.0.1960.0"  getversion = "Outlook 97 Ships only with Microsoft Exchange Server 5.5 and Microsoft Exchange 5.0 Service Pack 2"
        case "5.0.2178.0"  getversion = "Outlook 98 Initial release in Q1/1998."
        case "5.0.2819.0"  getversion = "Outlook 2000 Initial release in Q2/1999"
        case "5.0.3121.0"  getversion = "Outlook 2000 update included with Office 2000 SR1 (and SR1a)"
        case "5.0.3136.0"  getversion = "Outlook 2000 security update patch"
        case "5.0.3144.0"  getversion = "Outlook 2000 with Service Pack 2  installed"
        case "10.0.0.3311"  getversion = "Outlook 2002 with hotifx"
        case "8.00.3511"  getversion = "Outlook 97 Initial release in Q1/1997."
        case "8.01.3817"  getversion = "Outlook 97 Ships only with Microsoft Exchange 5.0 Service Pack 1"
        case "8.02.4212"  getversion = "Outlook 97 Ships only with Microsoft Office 97 SR-1"
        case "8.03.4629"  getversion = "Outlook 97 Ships only with Microsoft Exchange Server 5.5 and Microsoft Exchange 5.0 Service Pack 2"
        case "8.04.5619"  getversion = "Outlook 97 Ships only with Microsoft Office 97 SR-2"
        case "8.5.5104.6"  getversion = "Outlook 98 Initial release in Q1/1998. Is included with the Microsoft Exchange 5.5 Service Pack 1 CD-ROM (not available for download at the Microsoft site)"
        case "8.5.5603.0"  getversion = "Outlook 98 Microsoft Outlook 98 Security Patch 2 "
        case "8.5.7806"  getversion = "Outlook 98 Outlook 98 security update"
        case "9.0.0.2711"  getversion = "Outlook 2000 Initial Release in Q2/1999"
        case "9.0.0.3011"  getversion = "Outlook 2000 Microsoft Office 2000 Developer Edition released in Q3/1999"
        case "9.0.0.3821"  getversion = "Outlook 2000 update included with Office 2000 SR1 (and SR1a)"
        case "9.0.0.4201"  getversion = "Outlook 2000 security update patch installed."
        case "9.0.0.4527"  getversion = "Outlook 2000 with Service Pack 2  installed"
        case "9.0.0.6673"  getversion = "Outlook 2000 - SP3"
        case "10.0.0.2625"  getversion = "Outlook 2002 Initial release in Q1/2001."
        case "10.0.0.2627"  getversion = "Outlook 2002 Initial release in Q1/2001."
        case "10.0.0.3513"  getversion = "Outlook 2002 with Service Pack 1"
        case "10.0.0.3501"  getversion = "Outlook 2002 with Service Pack 1"
        case "10.0.0.4115"  getversion = "Outlook 2002 with Service Pack 2"
        case "10.0.0.3416"  getversion = "Outlook 2002 with Service Pack 1"
        case "10.0.0.4219"  getversion = "Outlook 2002 with Service Pack 2"
        case "10.0.0.6515"  getversion = "Outlook 2002 with Service Pack 3"
        case "10.0.0.6626"  getversion = "Outlook 2002 with Service Pack 3"
        case "11.0.5604.0"  getversion = "Outlook 2003 Initial release Oct. 2003 RTM Build"
        case "11.0.5606.0"  getversion = "Outlook 2003 Initial release Oct. 2003 RTM Build"
        case "11.0.5608.0"  getversion = "Outlook 2003 Initial release Oct. 2003 RTM Build"
        case "11.0.5703.0"  getversion = "Outlook 2003 with critical update releases on 11.4.2003"
        case "11.0.6353.0"  getversion = "Outlook 2003 SP1, released August 2003. "
        case "11.0.6352.0"  getversion = "Outlook 2003 SP1, released August 2003. "
	case "HTTP" getversion = "Outlook Web Access"
	case else getversion =  version
end select

end function

function getmode(clientmode)
	select case clientmode
		case 1 getmode = "Classic Online"
		case 2 getmode = "Cached Mode"
		case else getmode = " "
	end select
end function

public sub DisplayUsage
     WScript.echo "usage: cscript displogon.vbs <Servername> <Mode>"
     WScript.echo "Vaid Modes"
     WScript.echo "	OWA	- Show user logged on via Outlook Web Access"
     WScript.echo "	Outlook 	-Show users logged on via Outlook"
     WScript.echo "	Loggedon 	-Show All logged on Users"
     WScript.echo "	Loggedout	-Show Users not Logged On"
     WScript.echo "	ALL	-Show all users Logged in and Out"
end sub
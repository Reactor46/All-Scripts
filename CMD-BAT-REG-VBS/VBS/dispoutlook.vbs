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
           "  NEW adVarChar(255) AS OutlookVersion, " & _
           "  NEW adVarChar(255) AS OutlookFriendlyName, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS WMIDisplayName, " & _
           "      NEW adVarChar(255) AS WMILegacyDN, " & _
           "      NEW adVarChar(255) AS WMILoggedOnUserAccount, " & _
           "      NEW adVarChar(255) AS WMIClientVersion, " & _
           "      NEW adVarChar(255) AS WMIClientIP, " & _
           "      NEW adVarChar(255) AS WMIClientMode " & _
	   ")" & _
           "      RELATE OutlookVersion TO WMIClientVersion) AS rsOUWMI " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
insertOutlook(objParentRS)
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername &"/root/MicrosoftExchangeV2"
sqlstate = "Select * FROM Exchange_Logon where StoreType = 1 and ClientVersion <> 'SMTP' AND ClientVersion <> 'OLEDB'"
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_ExchangeLogons = objWMIExchange.ExecQuery(sqlstate,,48)
objChildRS.LockType = 3
Set objChildRS = objParentRS("rsOUWMI").Value
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
    Set objChildRS = objParentRS("rsOUWMI").Value 
    objChildRS.sort = "WMILoggedOnUserAccount"
    if objChildRS.recordcount <> 0 then
    cnum = 0 
    if showarg = "detail" then Wscript.echo objParentRS.fields("OutlookVersion") &  "	" & objParentRS.fields("OutlookFriendlyName") 		
    Do While Not objChildRS.EOF
	currec = objChildRS.fields("WMILoggedOnUserAccount") & objChildRS.fields("WMIClientVersion") & objChildRS.fields("WMIClientIP") & objChildRS.fields("WMIClientMode")
	if currec <> prevrec then 
		cnum = cnum + 1
		if showarg = "detail" then
			Wscript.echo "		" & objChildRS.fields("WMILoggedOnUserAccount") & "   " &  _
			getmode(objChildRS.fields("WMIClientMode")) & "  " & objChildRS("WMIClientIP") 
		end if
	end if
	prevrec = currec
	objChildRS.MoveNext
    Loop
    if showarg = "count" then Wscript.echo cnum & "	" & objParentRS.fields("OutlookVersion") &  "	" & objParentRS.fields("OutlookFriendlyName") 
    end if
    objParentRS.MoveNext
Loop

end sub

sub insertOutlook(objrecordset)

call insertoutlookdb(objrecordset,"4.0.994.0","Outlook 97 Initial release in Q1/1997.")
call insertoutlookdb(objrecordset,"5.0.1457.0","Outlook 97 Ships only with Microsoft Exchange 5.0 Service Pack 1")
call insertoutlookdb(objrecordset,"5.0.1458.0","Outlook 97 Ships only with Microsoft Office 97 SR-1")
call insertoutlookdb(objrecordset,"5.0.1960.0","Outlook 97 Ships only with Microsoft Exchange Server 5.5 and Microsoft Exchange 5.0 Service Pack 2")
call insertoutlookdb(objrecordset,"5.0.2178.0","Outlook 98 Initial release in Q1/1998.")
call insertoutlookdb(objrecordset,"5.0.2819.0","Outlook 2000 Initial release in Q2/1999")
call insertoutlookdb(objrecordset,"5.0.3121.0","Outlook 2000 update included with Office 2000 SR1 (and SR1a)")
call insertoutlookdb(objrecordset,"5.0.3136.0","Outlook 2000 security update patch")
call insertoutlookdb(objrecordset,"5.0.3144.0","Outlook 2000 with Service Pack 2  installed")
call insertoutlookdb(objrecordset,"10.0.0.3311","Outlook 2002 with hotifx")
call insertoutlookdb(objrecordset,"8.00.3511","Outlook 97 Initial release in Q1/1997.")
call insertoutlookdb(objrecordset,"8.01.3817","Outlook 97 Ships only with Microsoft Exchange 5.0 Service Pack 1")
call insertoutlookdb(objrecordset,"8.02.4212","Outlook 97 Ships only with Microsoft Office 97 SR-1")
call insertoutlookdb(objrecordset,"8.03.4629","Outlook 97 Ships only with Microsoft Exchange Server 5.5 and Microsoft Exchange 5.0 Service Pack 2")
call insertoutlookdb(objrecordset,"8.04.5619","Outlook 97 Ships only with Microsoft Office 97 SR-2")
call insertoutlookdb(objrecordset,"8.5.5104.6","Outlook 98 Initial release in Q1/1998. Is included with the Microsoft Exchange 5.5 Service Pack 1 CD-ROM (not available for download at the Microsoft site)")
call insertoutlookdb(objrecordset,"8.5.5603.0","Outlook 98 Microsoft Outlook 98 Security Patch 2 ")
call insertoutlookdb(objrecordset,"8.5.7806","Outlook 98 Outlook 98 security update")
call insertoutlookdb(objrecordset,"9.0.0.2711","Outlook 2000 Initial Release in Q2/1999")
call insertoutlookdb(objrecordset,"9.0.0.3011","Outlook 2000 Microsoft Office 2000 Developer Edition released in Q3/1999")
call insertoutlookdb(objrecordset,"9.0.0.3821","Outlook 2000 update included with Office 2000 SR1 (and SR1a)")
call insertoutlookdb(objrecordset,"9.0.0.4201","Outlook 2000 security update patch installed.")
call insertoutlookdb(objrecordset,"9.0.0.4527","Outlook 2000 with Service Pack 2  installed")
call insertoutlookdb(objrecordset,"9.0.0.6673","Outlook 2000 - SP3")
call insertoutlookdb(objrecordset,"10.0.0.2625","Outlook 2002 Initial release in Q1/2001.")
call insertoutlookdb(objrecordset,"10.0.0.2627","Outlook 2002 Initial release in Q1/2001.")
call insertoutlookdb(objrecordset,"10.0.0.3513","Outlook 2002 with Service Pack 1")
call insertoutlookdb(objrecordset,"10.0.0.3501","Outlook 2002 with Service Pack 1")
call insertoutlookdb(objrecordset,"10.0.0.4115","Outlook 2002 with Service Pack 2")
call insertoutlookdb(objrecordset,"10.0.0.3416","Outlook 2002 with Service Pack 1")
call insertoutlookdb(objrecordset,"10.0.0.4219","Outlook 2002 with Service Pack 2")
call insertoutlookdb(objrecordset,"10.0.0.6515","Outlook 2002 with Service Pack 3")
call insertoutlookdb(objrecordset,"10.0.0.6626","Outlook 2002 with Service Pack 3")
call insertoutlookdb(objrecordset,"11.0.5604.0","Outlook 2003 Initial release Oct. 2003 RTM Build")
call insertoutlookdb(objrecordset,"11.0.5606.0","Outlook 2003 Initial release Oct. 2003 RTM Build")
call insertoutlookdb(objrecordset,"11.0.5608.0","Outlook 2003 Initial release Oct. 2003 RTM Build")
call insertoutlookdb(objrecordset,"11.0.5703.0","Outlook 2003 with critical update releases on 11.4.2003")
call insertoutlookdb(objrecordset,"11.0.6353.0","Outlook 2003 SP1, released August 2003. ")
call insertoutlookdb(objrecordset,"11.0.6352.0","Outlook 2003 SP1, released August 2003. ")
call insertoutlookdb(objrecordset,"HTTP","Outlook Web Access")

end sub

sub insertoutlookdb(objParentRS,Outlookversion,Outlookfriendlyname)

objParentRS.addnew 
objParentRS("OutlookVersion") = Outlookversion
objParentRS("OutlookFriendlyName") = Outlookfriendlyname
objParentRS.update	

end sub

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
     WScript.echo "	Detail	- Show Details of exch Oulook version logged on"
     WScript.echo "	Count 	-Show just the count of Outlook versions logged on"
end sub
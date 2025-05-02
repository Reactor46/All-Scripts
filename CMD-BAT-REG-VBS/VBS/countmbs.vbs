wscript.echo countmb("myserver")

function countmb(servername)
mbnum = 0
totalmbs = 0
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set snrs = Com.Execute
mbQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & snrs.fields("distinguishedName") & "));name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		mbnum = 0
		rangeStep = 999
		lowRange = 0
		highRange = lowRange + rangeStep
		quit = false
		set objmailstore = getObject(objmailstorename)
		Do until quit = true
			on error resume next
			strCommandText = "homeMDBBL;range=" & lowRange & "-" & highRange
			objmailstore.GetInfoEx Array(strCommandText), 0
			if err.number <> 0 then quit = true
			varReports = objmailstore.GetEx("homeMDBBL")
			if quit <> true then mbnum = mbnum + ubound(varReports)+1
		        lowRange = highRange + 1
        		highRange = lowRange + rangeStep
		loop
		err.clear
		totalmbs = totalmbs + mbnum
		mbnum = 0
		Rs.MoveNext

Wend
countmb = totalmbs
end function
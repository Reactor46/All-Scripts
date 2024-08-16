set shell = createobject("wscript.shell")
Set objDictionary = CreateObject("Scripting.Dictionary")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDBPolicy);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Mailbox Stores Policies"
Wscript.echo
While Not Rs.EOF
		set objmailstorepolicy = getobject("LDAP://" & Rs.Fields("distinguishedName")) 
		pstring = ""
		if isarray(objmailstorepolicy.msExchPolicyOptionList) then
			tarray = objmailstorepolicy.GetEx("msExchPolicyOptionList")
			for each ent in tarray 
				if pstring <> "" then pstring = pstring & ","
				select case  Octenttohex(ent) 
					case "B202F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Limits"
					case "B102F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Database"
					case "B002F16F3FBAD211994500C04F79F1C9" pstring = pstring & "General"
					case "B302F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Full Text Indexing"
				end select
			next
		end if
		objDictionary.Add objmailstorepolicy.distinguishedName , dateadd("h",toffset,objmailstorepolicy.msExchPolicyLastAppliedTime) & "," & pstring
		wscript.echo "Policy Name : " & objmailstorepolicy.cn
		wscript.echo "Policy Last Applied : " & dateadd("h",toffset,objmailstorepolicy.msExchPolicyLastAppliedTime)
		wscript.echo "Policy Tabs Included : " & pstring
		wscript.echo
		wscript.echo "Servername,StoreName"
		if isarray(objmailstorepolicy.msExchPolicyListBL) then
			for each storeobj in objmailstorepolicy.msExchPolicyListBL
				set objmailstore = getobject("LDAP://" & storeobj) 
				strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
				wscript.echo strservername & "," & objmailstore.cn
			next
		else
			if objmailstorepolicy.msExchPolicyListBL <> "" then
				set objmailstore = getobject("LDAP://" & objmailstorepolicy.msExchPolicyListBL) 
				strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
				wscript.echo strservername & "," & objmailstore.cn
			else
				Wscript.echo "Policy Not Applied to any Stores"
			end if
		end if
		wscript.echo
		Rs.MoveNext

Wend
Rs.Close
set Rs = nothing
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPublicMDBPolicy);name,distinguishedName;subtree"
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Public Stores Policies"
Wscript.echo
While Not Rs.EOF
		set objmailstorepolicy = getobject("LDAP://" & Rs.Fields("distinguishedName")) 
		pstring = ""
		if isarray(objmailstorepolicy.msExchPolicyOptionList) then
			tarray = objmailstorepolicy.GetEx("msExchPolicyOptionList")
			for each ent in tarray 
				if pstring <> "" then pstring = pstring & ","
				select case  Octenttohex(ent) 
					case "A402F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Limits"
					case "A302F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Database"
					case "A102F16F3FBAD211994500C04F79F1C9" pstring = pstring & "General"
					case "A502F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Full Text Indexing"
					case "A202F16F3FBAD211994500C04F79F1C9" pstring = pstring & "Replication"
				end select
			next
		end if	
		objDictionary.Add objmailstorepolicy.distinguishedName , dateadd("h",toffset,objmailstorepolicy.msExchPolicyLastAppliedTime) & "," & pstring
		wscript.echo "Policy Name : " & objmailstorepolicy.cn
		wscript.echo "Policy Last Applied : " & dateadd("h",toffset,objmailstorepolicy.msExchPolicyLastAppliedTime)
		wscript.echo "Policy Tabs Included : " & pstring
		wscript.echo
		wscript.echo "Servername,StoreName"
		if isarray(objmailstorepolicy.msExchPolicyListBL) then
			for each storeobj in objmailstorepolicy.msExchPolicyListBL
				set objmailstore = getobject("LDAP://" & storeobj) 
				strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
				wscript.echo strservername & "," & objmailstore.cn
			next
		else
			if objmailstorepolicy.msExchPolicyListBL <> "" then
				set objmailstore = getobject("LDAP://" & objmailstorepolicy.msExchPolicyListBL) 
				strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
				wscript.echo strservername & "," & objmailstore.cn
			else
				Wscript.echo "Policy Not Applied to any Stores"
			end if
		end if
		wscript.echo
		Rs.MoveNext

Wend
Rs.Close
set Rs = nothing
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchExchangeServerPolicy);name,distinguishedName;subtree"
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Exchange Server Policies"
Wscript.echo
While Not Rs.EOF
		set objmailstorepolicy = getobject("LDAP://" & Rs.Fields("distinguishedName")) 
		
		wscript.echo "Policy Name : " & objmailstorepolicy.cn
		wscript.echo "Policy Last Applied : " & dateadd("h",toffset,objmailstorepolicy.msExchPolicyLastAppliedTime)
		wscript.echo
		wscript.echo "Servername"
		if isarray(objmailstorepolicy.msExchPolicyListBL) then
			for each storeobj in objmailstorepolicy.msExchPolicyListBL
				set objmailserver = getobject("LDAP://" & storeobj) 
				wscript.echo objmailserver.cn
			next
		else
			if objmailstorepolicy.msExchPolicyListBL <> "" then
				set objmailserver = getobject("LDAP://" & objmailstorepolicy.msExchPolicyListBL) 
				wscript.echo objmailserver.cn
			else
				Wscript.echo "Policy Not Applied to any Servers"
			end if
		end if
		wscript.echo
		Rs.MoveNext

Wend
Rs.Close
set Rs = nothing
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\exchpolices.csv",2,true)
wfile.writeline("""Servername"",""Storage Group"",""StoreName"",""Last Applied"",""Policy Tabs""")
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName;subtree"
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
		set objmailstore = getobject("LDAP://" & Rs.Fields("distinguishedName")) 
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		if isarray(objmailstore.msExchPolicyList) then
			tarray = objmailstore.GetEx("msExchPolicyList")
			for each ent in tarray 
				wfile.writeline strservername & "," & sgname & "," & objmailstore.cn & "," & objDictionary.Item(ent)
				
			next
		else
			if objmailstore.msExchPolicyList <> "" then
				wfile.writeline strservername & "," & sgname & "," & objmailstore.cn & "," & objDictionary.Item(objmailstore.msExchPolicyList)
			else
				wfile.writeline strservername  & "," & sgname & "," & objmailstore.cn & ",,No Polices"
			end if
		end if 
		Rs.MoveNext

Wend
Rs.Close
set Rs = nothing
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPublicMDB);name,distinguishedName;subtree"
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
		set objmailstore = getobject("LDAP://" & Rs.Fields("distinguishedName")) 
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		if isarray(objmailstore.msExchPolicyList) then
			tarray = objmailstore.GetEx("msExchPolicyList")
			for each ent in tarray 
				wfile.writeline strservername & "," & sgname & "," & objmailstore.cn & "," & objDictionary.Item(ent)
				
			next
		else
			if objmailstore.msExchPolicyList <> "" then
				wfile.writeline strservername & "," & sgname & "," & objmailstore.cn & "," & objDictionary.Item(objmailstore.msExchPolicyList)
			else
				wfile.writeline strservername  & "," & sgname & "," & objmailstore.cn & ",,No Polices"
			end if
		end if 
		Rs.MoveNext

Wend
Rs.Close


Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing


Function Octenttohex(OctenArry) 
ReDim aOut(UBound(OctenArry)) 
For i = 1 to UBound(OctenArry) + 1 
if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
else
aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
end if
Next 
Octenttohex = join(aOUt,"")
End Function




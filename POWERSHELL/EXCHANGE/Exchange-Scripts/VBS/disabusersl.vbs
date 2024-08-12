if wscript.arguments.length = 0 then 
	wscript.echo "Display Mode"
else
	if lcase(wscript.arguments(0)) = "remove" then 
		mode = "remove"
		wscript.echo "Remove Mode"
	else
		wscript.echo "Display Mode"
	end if
end if

Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\DisableUserReport.txt",8,true) 
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
           "  NEW adVarChar(255) AS GRPDisplayName, " & _
           "  NEW adVarChar(255) AS GRPDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS USDisplayName, " & _
           "      NEW adVarChar(255) AS USDN, " & _
           "      NEW adVarChar(255) AS USGRPDisplayName, " & _
           "      NEW adVarChar(255) AS USGRPDN " & _
	   ")" & _
           "      RELATE GRPDN TO USGRPDN) AS rsGRPUS " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
GALQueryFilter =  "(&(mailnickname=*)(|(objectCategory=group)))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
Com.ActiveConnection = Conn
Com.CommandText = strQuery
Set Rs = Com.Execute
while not rs.eof		
	objParentRS.addnew 
	objParentRS("GRPDisplayName") = rs.fields("displayname")
	objParentRS("GRPDN") = rs.fields("distinguishedName")
	objParentRS.update	
	rs.movenext
wend
GALQueryFilter = "(&(&(mailnickname=*)(objectCategory=person)(userAccountControl:1.2.840.113556.1.4.803:=2)))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
Com.ActiveConnection = Conn
Com.CommandText = strQuery
Set Rs1 = Com.Execute
Set objChildRS = objParentRS("rsGRPUS").Value
while not rs1.eof		
	if instr(rs1.fields("displayname"),"SystemMailbox{") = 0 then 
		set objuser = getobject("LDAP://" & replace(rs1.fields("distinguishedName"),"/","\/"))
		For each objgroup in objuser.groups
			objChildRS.addnew 
			objChildRS("USDisplayName") = rs1.fields("displayname")
			objChildRS("USDN") = rs1.fields("distinguishedName")
			objChildRS("USGRPDisplayName") = objgroup.name
			objChildRS("USGRPDN") = objgroup.distinguishedName
			objChildRS.update
		next
	end if
	rs1.movenext
wend
objParentRS.MoveFirst
wscript.echo "GroupName,Disabled User's Name"
wscript.echo
Do While Not objParentRS.EOF
     Set objChildRS = objParentRS("rsGRPUS").Value
     if objChildRS.recordCount <> 0 then
	Do While Not objChildRS.EOF
		Wscript.echo objParentRS.fields("GRPDisplayName") & "," & objChildRS.fields("USDisplayName")
		if mode = "remove" then 
			set objgroup = getobject("LDAP://" & replace(objChildRS.fields("USGRPDN"),"/","\/"))
			Set objUser = getobject("LDAP://" & replace(objChildRS.fields("USDN"),"/","\/"))
			objGroup.Remove(objUser.AdsPath) 
			objgroup.setinfo
			wscript.echo "User-Removed"
			wfile.writeline("User " & objUser.displayName & "(" & objUser.samaccountName & ") Removed from Group " & _
			objGroup.DisplayName & "(" & objGroup.distinguishedName & ")")
			if objUser.extensionAttribute10 = "" then 
				objUser.extensionAttribute10 = objgroup.distinguishedName
			else
				objUser.extensionAttribute10 = objUser.extensionAttribute10 & ";" & objgroup.distinguishedName
			end if
			objUser.setinfo
		end if
		objChildRS.MoveNext	
	loop
     end if 
     objParentRS.MoveNext
Loop
wfile.write report
wfile.close
set wfile = nothing
set fso = nothing
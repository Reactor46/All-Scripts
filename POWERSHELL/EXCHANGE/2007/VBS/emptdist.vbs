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
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
GALQueryFilter =  "(&(mail=*)(objectCategory=group))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
Com.ActiveConnection = Conn
Com.CommandText = strQuery
Com.Properties("SearchScope") = 2          ' we want to search everything
Com.Properties("Page Size") = 500          ' and we want our records in lots of 500 (must be < query limit)
Set Rs = Com.Execute
wscript.echo "# Members	GroupName"
wscript.echo
while not rs.eof
	set objgroup = getobject("LDAP://" & replace(rs.fields("distinguishedName"),"/","\/"))	
	numcheck = 0
	for each member in objgroup.members
		numcheck = numcheck + 1
		if numcheck = 10 then exit for
	next
	select case numcheck
		case 0 wscript.echo  "Empty" & "		" & rs.fields("displayname")
		       if mode = "remove" then
				contname = replace(objgroup.distinguishedName,"CN=" & objgroup.cn & ",","")
		       		wscript.echo
				wscript.echo "The Group " & rs.fields("displayname") & " in the " & contname & " container is Empty"
		       		WScript.StdOut.WriteLine "Do you wish to delete this List Press Y to Delete (This is a irreversible operation)"
		       		ans = WScript.StdIn.ReadLine 
		      		if lcase(ans) = "y" then
					set objcont =  getobject("LDAP://" & replace(contname,"/","\/"))
			 		objcont.delete "group", "CN="	& replace(objgroup.cn,"/","\/")
		          		objcont.setinfo
			  		Wscript.echo "Deleted Group " & objgroup.displayname
		      		else
			   		wscript.echo "Skipping"
		       		end if
		      		wscript.echo
			end if		
		case 10 wscript.echo "10+" & "		" & rs.fields("displayname")
		case else wscript.echo numcheck & "		" & rs.fields("displayname")
	end select
	rs.movenext
wend		
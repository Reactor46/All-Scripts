queryou = wscript.arguments(0)
report = "<table border=""1"" width=""100%"">" & vbcrlf
report = report & "  <tr>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">OU Name</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">OU Description</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">OU Path</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF""># Users</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF""># Contacts</font></b></td>" & vbcrlf
report = report & "</tr>" & vbcrlf
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
ouQuery = "<LDAP://" & ou & strNameingContext & ">;(&(objectCategory=organizationalUnit)(ou=" & queryou & "));description,name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = ouQuery
Set ouRs = Com.Execute
if ouRs.recordcount = 0 and queryou <> "rootdse"  Then
	wscript.echo "Cant find Ou"
Else
	If queryou = "rootdse"  Then
		ou = strNameingContext
	Else
		ou = ouRs.fields("distinguishedName")
	End if
	Wscript.echo "Users under " & ou
	wscript.echo
	uarray = ReportonOU("user")
	Wscript.echo
	Wscript.echo "Contacts under " & ou
	wscript.echo
	carray = ReportonOU("contact")

	For uic = LBound(uarray,3) To UBound(uarray,4)
		report = report & "<tr>" & vbcrlf
		report = report & "<td align=""center"">" & uarray(1,1,1,uic) & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & uarray(1,1,0,uic) & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & uarray(1,0,0,uic) & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & uarray(0,0,0,uic) & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & carray(0,0,0,uic) & "&nbsp;</td>" & vbcrlf
		report = report & "</tr>" & vbcrlf
	Next
	report = report & "</table>" & vbcrlf
	Set fso = CreateObject("Scripting.FileSystemObject")
	set wfile = fso.opentextfile("c:\temp\report.htm",2,true) 
	wfile.write report
	wfile.close
	set wfile = nothing
	set fso = nothing
End If


function ReportonOU(objecttype)

set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS SOADDisplayName, " & _
           "  NEW adVarChar(255) AS SOADDescription, " & _
           "  NEW adVarChar(255) AS SOADDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS MOADDisplayName, " & _
           "      NEW adVarChar(255) AS MOADDN " & ")" & _
           "      RELATE SOADDN TO MOADDN) AS rsSOMO " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1 

adsiQuery = "<LDAP://" & ou  & ">;(objectCategory=organizationalUnit);description,name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = adsiQuery
Set Rs = Com.Execute
While Not Rs.EOF
	objParentRS.addnew 
	objParentRS("SOADDisplayName") = rs.fields("name")
	If Not IsNull(rs.fields("description").value) then
		For Each val In rs.fields("description").value
			objParentRS("SOADDescription") = val
		Next
	Else
		objParentRS("SOADDescription") = " "
	End if
	objParentRS("SOADDN") = LCase(rs.fields("distinguishedName"))
	objParentRS.update	
	rs.movenext
Wend

objQuery = "<LDAP://" & ou  & ">;(&(mailnickname=*)(|(&(objectCategory=person)(objectClass=" & objecttype & "))));description,name,cn,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = objQuery
Set Rs1 = Com.Execute
Set objChildRS = objParentRS("rsSOMO").Value
While Not Rs1.EOF
	objChildRS.addnew 
	objChildRS("MOADDisplayName") = rs1.fields("name")
	objChildRS("MOADDN") = Replace(LCase(rs1.fields("distinguishedName")),"cn=" & LCase(rs1.fields("cn")) & ",","")
	objChildRS.update	
	rs1.movenext
Wend
anum = cInt(objParentRS.recordcount)
reDim objcntarray(1,1,1,anum)
i = 0
objParentRS.MoveFirst
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("rsSOMO").Value 
	wscript.echo objParentRS("SOADDisplayName") & "	- " & objParentRS("SOADDescription") &  " : "& objChildRS.recordCount
    objcntarray(1,1,1,i) =  objParentRS("SOADDisplayName")
    objcntarray(1,1,0,i) =  objParentRS("SOADDescription") 
    objcntarray(1,0,0,i) =  objParentRS("SOADDN") 
    objcntarray(0,0,0,i) =  objChildRS.recordCount
    objParentRS.movenext
	i = i + 1
Loop
ReportonOU = objcntarray
objParentRS.close
objChildRS.close
rs.close
rs1.close
Set objParentRS = Nothing
Set objChildRS = Nothing
set conn1 = Nothing
Set Rs1 = Nothing
Set rs = Nothing

End function





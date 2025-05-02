rem on error resume next
set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\permissionDump.csv",2,true)
wfile.close
set wfile = nothing
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString	
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Query = "<LDAP://" & strNameingContext & ">;(&(&(&(& (mailnickname=*)" & _
	"(!msExchHideFromAddressLists=TRUE) (| (&(objectCategory=person)(objectClass=user)" & _
	"(|(homeMDB=*)(msExchHomeServerName=*))) )))))" & _
	";samaccountname,legacyExchangeDN,msExchHomeServerName,displayname,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = Query
Com.Properties("Page Size") = 1000

set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS UOADDisplayName, " & _
           "  NEW adVarChar(255) AS UOADTrusteeName, " & _
           "  NEW adVarChar(255) AS UOADLegacyDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS MRmbox, " & _
           "      NEW adVarChar(255) AS MRFolder, " & _
           "      NEW adVarChar(255) AS MRTrusteeName, " & _
           "      NEW adVarChar(255) AS MRRights) " & _
           "      RELATE UOADLegacyDN TO MRTrusteeName) AS rsUOMR" 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1

Set Rs = Com.Execute
While Not Rs.EOF
	inplinearray = Split(rs.fields("msExchHomeServerName").value, "=", -1, 1)
	cmdexe = "c:\windows\system32\cscript.exe c:\temp\dumpperms2.vbs " & rs.fields("samaccountname").value & " " & inplinearray(ubound(inplinearray))     
	wscript.echo cmdexe           
	ef =  WshShell.run(cmdexe,1,true)
	wscript.echo rs.fields("samaccountname").value
	objParentRS.addnew 
	objParentRS("UOADDisplayName") = rs.fields("displayname")
	objParentRS("UOADTrusteeName") = rs.fields("samaccountname")
	objParentRS("UOADLegacyDN") = rs.fields("legacyExchangeDN")
	objParentRS.update
	rs.movenext
Wend
Set objChildRS = objParentRS("rsUOMR").Value
set wfile = fso.opentextfile("c:\temp\permissionDump.csv",1,true)
Do While Not wFile.AtEndOfStream
	impline = wfile.readline
	splitarray = split(impline,",")
	objChildRS.addnew
	objChildRS("MRmbox") = splitarray(1)
	objChildRS("MRFolder") = splitarray(2)
	objChildRS("MRTrusteeName") = splitarray(3)
	objChildRS("MRRights") = splitarray(4)
	objChildRS.update
loop

objParentRS.MoveFirst
Do While Not objParentRS.EOF
	Set objChildRS = objParentRS("rsUOMR").Value
	if objChildRS.recordcount <> 0 then wscript.echo objParentRS("UOADDisplayName")
	Do While Not objChildRS.EOF
		wscript.echo "   " & objChildRS.fields("MRmbox") & _
			" "  & objChildRS.fields("MRFolder") & " " &  objChildRS.fields("MRRights")
		objChildRS.movenext
	loop
	objParentRS.MoveNext
loop

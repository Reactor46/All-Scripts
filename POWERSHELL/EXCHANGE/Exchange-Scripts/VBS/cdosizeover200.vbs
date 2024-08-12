Public Const CdoDefaultFolderCalendar = 0
servername = wscript.arguments(0)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("C:\" & servername & ".csv",2,true)
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
wscript.echo "Mailboxes Over 200 MB"
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		msize = getsize(servername,rs1.fields("mailnickname"))
		if not isnull(msize) then
			if msize > 200 then 
				wscript.echo rs1.fields("mailnickname") & " " & msize
				wfile.writeline rs1.fields("mailnickname") & "," & replace(msize,",","")
			end if
		end if
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
wfile.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"


function getsize(servername,mailboxname)
on error resume next
Set objSession = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set oInfoStores = objSession.InfoStores
For Each oInfoStore In oInfoStores
      If InStr(1, oInfoStore.Name, "Mailbox - ", 1) <> 0 Then
         getsize = formatnumber((oInfoStore.Fields(&He080014)/1048576),2)
      End If
Next
End function


action = wscript.arguments(0)
servername = wscript.arguments(1)
foldername = wscript.arguments(2)
public admnamefold 
admnamefold = "http://" & servername & "/ExAdmin/Admin/domain.com/public folders/" &  foldername & "/"
Set objX = CreateObject("Microsoft.XMLHTTP")
objX.Open "PROPFIND", admnamefold, FALSE, "", ""
strR = "<?xml version='1.0'?>"
strR = strR & "<a:propfind xmlns:a='DAV:' xmlns:e='http://schemas.microsoft.com/mapi/proptag/'>"
strR = strR & "<a:prop><e:x66980102/></a:prop></a:propfind>"
objX.SetRequestHeader "Content-type:", "text/xml"
objX.SetRequestHeader "Depth", "0"
objX.send(strR)
 
set docback = objX.responseXML
Dim objNodeList
Set objNodeList = docback.getElementsByTagName("d:x66980102")
For i = 0 TO (objNodeList.length -1)
  Set objNode = objNodeList.nextNode
  strrep = Octenttostring(objNode.nodeTypedValue)
  call Repaction(strrep,action,servername)
Next


sub Repaction(strrep,action,servername)

select case action
	case "display" wscript.echo "Current Replica's"
		       wscript.echo
		       wscript.echo strrep
	case "add"  call addrep(strrep,servername)
	case "remove" call deleterep(strrep,servername)
	case else Wscript.echo "Invalid Commmand use add,remove or display"
end select

end sub

sub addrep(strrep,servername)
aexist = 0
reparray = split(strrep,chr(10),-1,1)
pubfolderLDN = PublicfolderLDN(Servername)
for i = lbound(reparray) to ubound(reparray)
	if ucase(reparray(i)) = ucase(pubfolderLDN) then aexist = 1
next	
if aexist = 1 then 
	wscript.echo "Replica already Exist for this server"
	call Repaction(strrep,"display",servername)
else
	strrep = replace(strrep,chr(10),chr(0)) & pubfolderLDN & chr(0)
	call proppatchrep(strrep)
end if
end sub

sub deleterep(strrep,servername)
aexist = 0
nstr = ""
reparray = split(strrep,chr(10),-1,1)
pubfolderLDN = PublicfolderLDN(Servername)
if ubound(reparray) > 1 then
	for i = lbound(reparray) to ubound(reparray)-1
		if ucase(reparray(i)) = ucase(pubfolderLDN) then 
			aexist = 1
		else
			nstr = reparray(i) & chr(0)
		end if
	next
else
	Wscript.echo "You cant delete the only replica of this folder !"
end if	
if aexist = 1 then 
	call proppatchrep(nstr)
else
	wscript.echo "NO Replica Exists for this server"
	call Repaction(strrep,"display",servername)
end if
end sub

sub proppatchrep(strrep)

set convobj = CreateObject("Msxml2.DOMDocument.4.0")
Set oRoot = convobj.createElement("test")
oRoot.dataType = "bin.base64"
OriginalLocale = SetLocale(1033)
set stm =  CreateObject("ADODB.Stream") 
stm.Type = 2 
stm.Charset = "x-ansi" 
stm.Open 
stm.WriteText strrep 
stm.Position = 0 
stm.Type = 1 
oRoot.nodeTypedValue = stm.Read
SetLocale(OriginalLocale)
xmlstring = "<?xml version=""1.0"" encoding=""utf-8""?>"
xmlstring = xmlstring & "<d:propertyupdate xmlns:d=""DAV:"" xmlns:m=""http://schemas.microsoft.com/mapi/proptag/""" _
& " xmlns:b=""urn:schemas-microsoft-com:datatypes/"">" & vbcrlf & "<d:set>" & vbcrlf & "<d:prop>" & vbcrlf & "<m:0x66980102 b:dt=""bin.base64"">"
xmlstring = xmlstring & oRoot.text & vbcrlf
xmlstring = xmlstring & "</m:0x66980102>" & vbcrlf & "</d:prop>" & vbcrlf & "</d:set>" & vbcrlf & "</d:propertyupdate>"
Set xmlReq = CreateObject("Microsoft.XMLHTTP")
xmlReq.open "PROPPATCH",admnamefold, False
xmlReq.setRequestHeader "Content-Type", "text/xml"
xmlReq.send xmlstring
If xmlReq.status >= 500 Then
   wscript.echo "Status: " & xmlReq.status
   wscript.echo "Status text: An error occurred on the server."
ElseIf xmlReq.status = 207 Then
   wscript.echo "Replica Set sussessfully"
Else
   wscript.echo "Status: " & xmlReq.status
   wscript.echo "Status text: " & xmlReq.statustext
   wscript.echo "Response text: " & xmlReq.responsetext
End If

end sub

Function Octenttostring(OctenArry) 
  ReDim aOut(UBound(OctenArry)) 
  For i = 1 to UBound(OctenArry) + 1 
	if ascb(midb(OctenArry,i,1)) <> 0 then 
		aOut(i-1) = chr(ascb(midb(OctenArry,i,1)))
	else	
		aOut(i-1) = chr(10)
	end if
  Next 
  Octenttostring = join(aOUt,"")
End Function 

function PublicfolderLDN(Servername)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
set conn1 = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	defpf =  "(&(objectCategory=msExchPFTree)(msExchPFTreeType=1))"
	strQuery = "<LDAP://"  & strNameingContext & ">;" & defpf  & ";distinguishedName;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not rs1.eof
		snameq = "(&(&(objectCategory=msExchPublicMDB)(msExchOwningPFTree=" & rs1.fields("distinguishedName") _
		& ")(msExchOwningServer=" & rs.fields("distinguishedName") & ")))"
		strQuery = "<LDAP://"  & strNameingContext & ">;" & snameq  & ";distinguishedName,legacyExchangeDN;subtree"
		com.Properties("Page Size") = 100
		Com.CommandText = strQuery
		Set Rs2 = Com.Execute
		while not rs2.eof
			PublicfolderLDN = rs2.fields("legacyExchangeDN")
			rs2.movenext
		wend
		rs1.movenext
	wend	
	rs.movenext
wend
end function

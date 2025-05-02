ScpUrlGuidString = "77378F46-2C66-4aa9-A6A6-3E7A48B19596"
ScpPtrGuidString = "67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68"

set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://CN=Microsoft Exchange,CN=Services," & strNameingContext & ">;(&(objectClass=serviceConnectionPoint)"  _
	& "(|(keywords=" & ScpPtrGuidString & ")(keywords=" & ScpUrlGuidString & ")));cn,name,serviceBindingInformation,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof
	wscript.echo rs.fields("cn")
	call queryautodiscovery(wscript.arguments(0),rs.fields("serviceBindingInformation").Value)
	rs.movenext
wend	

sub queryautodiscovery(emailaddress,casAddress)
wscript.echo "Using AutoDisover Address : " & casAddress(0)
autodiscoResponse = "<Autodiscover xmlns=""http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"">" _
& "	<Request>" _
& "		<EMailAddress>" + emailaddress + "</EMailAddress>" _
& "		<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>" _
& "	</Request>" _
& "</Autodiscover>"
set req = createobject("MSXML2.ServerXMLHTTP.6.0")
req.Open "Post",casAddress(0) ,False
req.SetOption 2, 13056 
req.setRequestHeader "Content-Type", "text/xml"
req.setRequestHeader "Content-Length", len(autodiscoResponse)
req.send autodiscoResponse 
wscript.echo req.responsetext



end sub
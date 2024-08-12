Dim Rec,Rs,strURLInbox,msgobj,msgobj1,flds,objArgs,strView
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")


GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=/o=yourorg/ou=First Administrative Group/cn=Configuration/cn=Servers/cn=servername)) )))))"
strQuery = "<LDAP://" & strDefaultNamingContext & ">;" & GALQueryFilter & ";samaccountname,mail;subtree"

Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery
oComm.Properties("Sort on") = "givenname"

Set rs = oComm.Execute
while not rs.eof
 oofstat = getoofstat(rs.fields("mail"))   
 wscript.echo rs.fields("Samaccountname") & " " & oofstat             
 rs.movenext
wend


function getoofstat(emailaddress)
on error resume next
Set Person = CreateObject("CDO.Person")
strURL = "mailto:" & emailaddress
Person.DataSource.Open strURL
Set Mailbox = Person.GetInterface("IMailbox")
inbstr = Mailbox.basefolder & "/non_ipm_subtree/" 
Set Rec = CreateObject("ADODB.Record")
Rec.Open inbstr
getoofstat = rec.fields("http://schemas.microsoft.com/exchange/oof-state")
set rec = nothing

end function

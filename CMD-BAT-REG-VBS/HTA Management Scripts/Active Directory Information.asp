<%@ Language=VBScript %>
<html>
<head>
<title>Phone and Email List</title>
</head>
<body bgcolor="#FFFFFF">
<%
'=========Account and connection string information for LDAP=======
Set objDomain = GetObject ("GC://RootDSE")
objADsPath = objDomain.Get("defaultNamingContext")
Set objDomain = Nothing
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.provider ="ADsDSOObject"
objConn.Properties("User ID") = "domain\username" 'domain account with read access to LDAP
objConn.Properties("Password") = "password" 'domain account password
objConn.Properties("Encrypt Password") = True
objConn.open "Active Directory Provider"
Set objCom = CreateObject("ADODB.Command")
Set objCom.ActiveConnection = objConn
objCom.CommandText ="select name,telephonenumber,mobile,mail,company,title,department,sAMAccountName,sn,userAccountControl,msexchhidefromaddresslists FROM 'GC://"+objADsPath+"' where sAMAccountname='*' ORDER by sAMAccountname"

'=======Executre queury on LDAP for all accounts=========
Set objRS = objCom.Execute

'Loop through records and write out all information using ASP
Response.Write "<center><table>"

Do While Not objRS.EOF Or objRS.BOF
Response.Write "<tr>"
Response.Write "<td>"
Response.Write objRS("name")
Response.Write "</td><td>"
Response.Write objRS("mail")
Response.Write "</td><td>"
Response.Write objRS("telephonenumber")
Response.Write "</td><td>"
Response.Write objRS("mobile")
Response.Write "</td><td>"
Response.Write objRS("company")
Response.Write "</td><td>"
Response.Write objRS("department")
Response.Write "</td><td>"
Response.Write objRS("title")
Response.Write "</td><td>"
Response.Write objRS("userAccountControl")
Response.Write "</td><td>"
Response.Write objRS("sAMAccountName")
Response.Write "</td><td>"
Response.Write objRS("msexchhidefromaddresslists")
Response.Write "</td><td>"
Response.Write FormatDateTime(Date,2)
Response.Write "</td>"
Response.Write "</tr>"
objRS.MoveNext
Response.Flush
Loop

Response.Write "</table>"

'Close objects and remove from memory
objRS.Close
objConn.Close
Set objRS = Nothing
Set objConn = Nothing
Set objCom = Nothing
Set objADsPath = Nothing
Set objDomain = Nothing
%>
</body>
</html> 
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Resubmit</title>
</head>

<body>
<%
dim rec,oCon,Href,msgobj,resub,Subject,Toaddress,pickupdirectory,rfcmsg
pickupdirectory = "c:\Program Files\Exchsrvr\Mailroot\vsi 1\PickUp"
Resub = request.form("resub")
Subject = request.form("Subject")
Toaddress = request.form("Toaddress")
set rec = createobject("ADODB.Record")
Set oCon = CreateObject("ADODB.Connection")
Href = request.querystring("Href")
oCon.ConnectionString = Href
oCon.Provider = "ExOledb.Datasource"
oCon.Open
set msgobj = createobject("CDO.Message")
msgobj.datasource.open Href,oCon,3
if Resub = "Yes" then
	response.write "Message Resubmitted to :" & Toaddress
	response.write "<br><bR><a href=""resubmain.asp"">Back To Main page</a></p>"
	msgobj.fields("urn:schemas:mailheader:subject") = Subject
	msgobj.fields.update
	set stm = msgobj.getstream
	stm.type = 2
	stm.Charset = "x-ansi"
	rfcmsg = stm.readtext
	rfcmsg = "x-sender: " & msgobj.fields("urn:schemas:httpmail:fromemail") & vbcrlf & rfcmsg
	rfcmsg = "x-receiver: " & Toaddress & vbcrlf & rfcmsg
	stm.position = 0 
	stm.writetext = rfcmsg
	Randomize   ' Initialize random-number generator.
	rndval = Int((20000000000 * Rnd) + 1)   
	stm.savetofile pickupdirectory & "\" & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
else
	response.write "<form method=""POST""><p>Check To Email address and Subject and select Submit to Resend Message</p>"
	response.write "<p>Message To :<input type=""text"" name=""Toaddress"" size=""80""></p>"
	response.write "<p>Subject :<input type=""text"" name=""Subject"" size=""80"" value=""" & msgobj.subject & """></p>"
	response.write "<p><input type=""submit"" value=""Submit"" action=""expresub.asp"" name=""B1""><input type=""reset"" value=""Reset"" name=""B2""></p>"
	response.write "<input type=""hidden"" name=""Resub"" value=""Yes"">"
	response.write "</form><p>Message Body</p><p>&nbsp;</p></form>"
	response.write msgobj.htmlbody
	set msgobj = nothing
end if
%>
</body>

</html>
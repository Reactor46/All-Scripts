

<%
dim server,conn,Exp,mailbox,strURL,Href,strQuery,req,oNode1,onode2,oNode,oNodeList,oResponseDoc,oclasslist,rec,msgobj,stm,rs

set req = createobject("microsoft.xmlhttp")
set rec = createobject("ADODB.Record")
Set oCon = CreateObject("ADODB.Connection")
Href = request.querystring("Href")
if Href = "" then Href = request.form("Href")
Exp = request.querystring("exp")
response.write "<form method=""POST""><p><b>Enter Folder URL (eg Http://servername/public/folder)</b></p>"
response.write "<p><input type=""text"" name=""HREF"" size=""100""><input type=""submit"" value=""Submit"" action=""resubmain.asp"" name=""B1"">"
response.write "<input type=""reset"" value=""Reset"" name=""B2""></p></form>"
if Href = "" then

else
	if Exp = "Yes" then
			oCon.ConnectionString = Href
			oCon.Provider = "ExOledb.Datasource"
			oCon.Open
			set msgobj = createobject("CDO.Message")
			msgobj.datasource.open Href,oCon,3
			set stm = msgobj.getstream
			stm.type = 1
			response.clear
			response.Contenttype= msgobj.fields("urn:schemas:mailheader:content-type").value
			Response.AddHeader "Content-Disposition","attachment;filename=export.txt"
			Response.AddHeader "Content-length",stm.size
			response.binarywrite stm.read
			Response.End 
			rem response.write msgobj.fields("urn:schemas:httpmail:htmldescription").value
        else
			response.write "<table border=""1"" width=""100%"" cellspacing=""1""><tr><td align=""center""><font face=""Courier""><b>Date</b></font></td>"
			response.write "<td align=""center""><font face=""Courier""><b>From </b></font></td><td align=""center""><font face=""Courier""><b>Subject</b></font></td>"
			response.write "<td align=""center""><font face=""Courier""><b>Export</b></font></td>"
			response.write "<td align=""center""><font face=""Courier""><b>Resub</b></font></td></tr>"
			strQuery = "SELECT ""DAV:href"", ""urn:schemas:mailheader:subject"", ""urn:schemas:httpmail:fromname"", "
			strQuery = strQuery & """urn:schemas:httpmail:datereceived"" FROM scope('shallow traversal of """
			strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False "
			oCon.ConnectionString = Href
			oCon.Provider = "ExOledb.Datasource"
			oCon.Open
			Set Rs = CreateObject("ADODB.Recordset")
			rs.open strQuery,oCon
			while not rs.eof
				response.write "<tr><td align=""center"">" & rs.fields("urn:schemas:httpmail:datereceived").value & "</td>" & vbcrlf
				response.write "<td align=""left"">" & left( rs.fields("urn:schemas:httpmail:fromname").value,25) & "</td>" & vbcrlf
				response.write "<td align=""left"">" & rs.fields("urn:schemas:mailheader:subject").value & "</td>" & vbcrlf
				response.write "<td align=""left"">" & "<a href=""resubmain.asp?Exp=Yes&Href=" & rs.fields("DAV:href").value & """>Export</a>" & "</td>" & vbcrlf
				response.write "<td align=""left"">" & "<a href=""expresub.asp?Exp=Yes&Href=" & rs.fields("DAV:href").value & """>Resubmit</a>" & "</td></tr>" & vbcrlf
				rs.movenext
			wend
	end if
End if


%>

</table>
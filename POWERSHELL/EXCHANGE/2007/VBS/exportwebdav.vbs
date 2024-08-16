server = "servername"
mailbox = "mailbox"
set fso = createobject("Scripting.FileSystemObject")
strURL = "http://" & server & "/exchange/" & mailbox & "/inbox/"
strURL1 = "http://" & server & "/exchange/" & mailbox & "/sent items/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"", ""urn:schemas:httpmail:subject"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """DAV:contentclass"" = 'urn:content-classes:message'</D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
rem req.setrequestheader "Range", "rows=0-100"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	df = expmail(oNode.Text,oNode1.text)
   Next	
Else
End If

function expmail(displayname,href)
req.open "GET", href, false
req.setRequestHeader "Translate","f"
req.send
fname = replace(replace(replace(replace(replace(displayname,":","-"),"\",""),"/",""),"?",""),chr(34),"")
fname = replace(replace(replace(replace(replace(replace(fname,"<",""),">",""),chr(11),""),"*",""),"|",""),"(","")
fname = replace(replace(replace(fname,")",""),chr(12),""),chr(15),"")
fname = "c:\exp\" & fname
wscript.echo fname
set stm = createobject("ADODB.Stream")
stm.open
msgstring = req.responsetext
stm.type = 2
stm.Charset = "x-ansi"
stm.writetext msgstring,0
stm.Position = 0
stm.type = 1
stm.savetofile fname
set stm = nothing

end function

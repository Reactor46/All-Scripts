Href = "http://servername/public/foldername"
sourceaddress = "blah@domain.com"
targetaddress = "traget@domain.com"
smtpservername = "servername"

set rec = createobject("ADODB.Record")
Set oCon = CreateObject("ADODB.Connection")
Set fso = CreateObject("Scripting.FileSystemObject")
oCon.ConnectionString = Href
oCon.Provider = "ExOledb.Datasource"
oCon.Open
strQuery = "SELECT ""DAV:href"", ""urn:schemas:mailheader:subject"", ""urn:schemas:httpmail:fromname"", "
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" FROM scope('shallow traversal of """
strQuery = strQuery & Href & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False "
Set Rs = CreateObject("ADODB.Recordset")
Rs.CursorLocation = 2 'adUseServer = 2, adUseClient = 3
rs.open strQuery,oCon,3
while not rs.eof
	sendmes(rs.fields("DAV:href"))
	rs.delete 1
	rs.movenext
wend


function sendmes(forwardhref)

set imsg1 = CreateObject("CDO.Message")
imsg1.datasource.open forwardhref
set stm1 = imsg1.getstream
Randomize   ' Initialize random-number generator.
rndval = Int((20000000000 * Rnd) + 1)   
fname = "c:\temp\" & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
stm1.savetofile fname

Set iMsg = CreateObject("CDO.Message")

With iMsg
   .From    = sourceaddress 
   .To      = targetaddress
   .Subject = "Forwarded message"
End With
iMsg.AddAttachment fname

set file = fso.getfile(fname)
file.delete
set file = nothing

iMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
iMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpservername
iMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
iMsg.Configuration.Fields.Update
iMsg.Send
wscript.echo "Message Sent"

end function


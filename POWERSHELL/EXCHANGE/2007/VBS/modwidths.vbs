servername = "servername"
mailbox = "mailbox"

xmlhdr = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
xmlhdr = xmlhdr & "<D:propertyupdate xmlns:D=""DAV:"" xmlns:a=""http://schemas.microsoft.com/exchange/"">"
xmlhdr = xmlhdr &  "<D:set>"
xmlhdr = xmlhdr &  "<D:prop>" 
xmlftr = xmlftr &  "</D:prop>" 
xmlftr = xmlftr &  "</D:set>" 
xmlftr = xmlftr &  "</D:propertyupdate>" 

sDestinationURL = "http://" & servername & "/exchange/" & mailbox & "/"
xmlstr = xmlhdr &  "<a:webclientnavbarwidth>200</a:webclientnavbarwidth>" & xmlftr
df = sendrequest(xmlstr,sDestinationURL)

sDestinationURL = "http://" & servername & "/exchange/" & mailbox & "/inbox"
xmlstr = xmlhdr &  "<a:wcviewwidth>400</a:wcviewwidth>"  & xmlftr
df = sendrequest(xmlstr,sDestinationURL)

function sendrequest(xmlstr,sDestinationURL)
Set XMLreq = CreateObject("Microsoft.xmlhttp")
XMLreq.open "PROPPATCH", sDestinationURL, False
XMLreq.setRequestHeader "Content-Type", "text/xml;"
XMLreq.setRequestHeader "Translate", "f"
XMLreq.setRequestHeader "Content-Length:", Len(xmlstr)
XMLreq.send(xmlstr)

If (XMLreq.Status >= 200 And XMLreq.Status < 300) Then
  Wscript.echo "Success!   " & "Results = " & XMLreq.Status & ": " & XMLreq.statusText
ElseIf XMLreq.Status = 401 then
  Wscript.echo "You don't have permission to do the job! Please check your permissions on this item."
Else
  Wscript.echo "Request Failed.  Results = " & XMLreq.Status & ": " & XMLreq.statusText
End If
set XMLreq = nothing
end function



OWAURL = "http://server/public/folder"
set xmlobjreq = createobject("Microsoft.XMLHTTP")
xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:e='http://schemas.microsoft.com/mapi/proptag/'><a:prop><e:x0FFF0102/></a:prop></a:propfind>"
xmlobjreq.open "PROPFIND", OWAURL, false, "", ""
xmlobjreq.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
xmlobjreq.setRequestHeader "Depth", "0"
xmlobjreq.setRequestHeader "Translate", "f"
xmlobjreq.send xmlreqtxt
set oResponseDoc = xmlobjreq.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("d:x0FFF0102")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	wscript.echo Octenttohex(oNode.nodeTypedValue)
Next

Function Octenttohex(OctenArry)  
  ReDim aOut(UBound(OctenArry)) 
  For i = 1 to UBound(OctenArry) + 1 
    if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
    	aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
    else
	aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
    end if
  Next 
  Octenttohex = join(aOUt,"")
End Function 
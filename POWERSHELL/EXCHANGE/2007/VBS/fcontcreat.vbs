set req = createobject("microsoft.xmlhttp")
folderurl = "http://servername/public/folder"

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:s='http://schemas.microsoft.com/exchange/security/'><a:prop><s:creator/></a:prop></a:propfind>"
req.open "PROPFIND", folderurl, false, "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("S:nt4_compatible_name")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	wscript.echo "Folder Created By : " & oNode.text
   next
end if

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:s='http://schemas.microsoft.com/exchange/security/'><a:prop><s:descriptor/></a:prop></a:propfind>"
req.open "PROPFIND", folderurl, false, "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
set oResponseDoc = req.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("S:effective_aces")
set oNode = oNodeList.nextnode
set oNodeList1 = oNode.selectnodes("S:access_allowed_ace/S:sid/S:nt4_compatible_name")
set oNodeList2 = oNode.selectnodes("S:access_allowed_ace/S:access_mask")
For nl = 1 To oNodeList1.length
	set oNode1 = oNodeList1.nextnode
	set oNode2 = oNodeList2.nextnode
	binmask = getbinval(oNode2.Text)
	if len(binmask) > 16 then
		if mid(right(binmask,16),1,1) = 1 then
			Contacts = Contacts & oNode1.Text & ","
		end if
	end if
Next
set oNodeList3 = oNode.selectnodes("S:access_denied_ace/S:sid/S:nt4_compatible_name")
set oNodeList4 = oNode.selectnodes("S:access_denied_ace/S:access_mask")
For nl1 = 1 To oNodeList3.length
	set oNode3 = oNodeList3.nextnode
	set oNode4 = oNodeList4.nextnode
	binmask = getbinval(oNode4.Text)
	if len(binmask) > 16 then
		if mid(right(binmask,16),1,1) = 1 then
			Contacts = replace(Contacts,oNode3.Text & ",","")
		end if
	end if
Next
wscript.echo "Folder Contacts : " & Contacts



function getbinval(mask)
binval = " "
for bv = 1 to len(mask)
	select case mid(mask,bv,1)
		case "f" binval = binval & "1111"
		case "e" binval = binval & "1110"
		case "d" binval = binval & "1101"
		case "c" binval = binval & "1100"
		case "b" binval = binval & "1011"
		case "a" binval = binval & "1010"
		case "9" binval = binval & "1001"
		case "8" binval = binval & "1000"
		case "7" binval = binval & "0111"
		case "6" binval = binval & "0110"
		case "5" binval = binval & "0101"
		case "4" binval = binval & "0100"
		case "3" binval = binval & "0011"
		case "2" binval = binval & "0010"
		case "1" binval = binval & "0001"
		case "0" binval = binval & "0000"
	end select
next
getbinval = binval
end function
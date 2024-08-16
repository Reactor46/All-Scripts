shareDN = "temp"
ShareHostName = "servername"
Shareuf = "6"
ShareURI = "file://servername/temp"

snServername = "mailserver"
mbMailboxName = "mailbox"

Const PR_PARENT_ENTRYID = &H0E090102
set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
xdXmlDocument.async="false"
ifound = false
Set objSession   = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,snServername & vbLF & mbMailboxName
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
set non_ipm_rootfolder = objSession.getfolder(CdoFolderRoot.fields.item(PR_PARENT_ENTRYID),CdoInfoStore.id)
For Each soStorageItem in non_ipm_rootfolder.HiddenMessages 
	If soStorageItem.Type = "IPM.Configuration.Owa.DocumentLibraryFavorites" Then
		ifound = true 
		Set actionItem = soStorageItem
	End if
Next 
If ifound = false Then
	wscript.echo "No Storage Item Found"
Else
	On Error Resume Next 
	hexString = actionItem.fields(&h7C080102).Value
	If Err.number <> 0 Then
		On Error goto 0
		wscript.echo "Property not set"
		actionItem.fields.Add &h7C080102, vbBlob 
		actionItem.fields(&h7C080102).Value = StrToHexStr("<docLibs></docLibs>")
		actionItem.update
		hexString = actionItem.fields(&h7C080102).Value
	End If
	On Error goto 0
	wscript.echo hextotext(hexString)
	xdXmlDocument.loadxml(hextotext(hexString))
	Set xnNodes = xdXmlDocument.selectNodes("//docLibs")
	update = false
	Call Adddoclib(shareDN,ShareHostName,Shareuf,ShareURI,xnNodes)
	If update = True Then 
		nval = StrToHexStr(CStr(xdXmlDocument.xml))
		actionItem.fields(&h7C080102).Value = nval
		actionItem.update
		wscript.echo "Storage Object Updated"
	Else
		wscript.echo "No Updates Performed"
	End If
End If



Function hextotext(binprop)
arrnum = len(binprop)/2
redim aout(arrnum)
slen = 1
for i = 1 to arrnum
	if CLng("&H" & mid(binprop,slen,2)) <> 0 then
		aOut(i) = chr(CLng("&H" & mid(binprop,slen,2)))		
	end if
	slen = slen+2
next
hextotext = join(aOUt,"")
end Function

Function StrToHexStr(strText) 
 Dim i, strTemp 
 For i = 1 To Len(strText) 
  strTemp = strTemp & Right("0" & Hex(Asc(Mid(strText, i, 1))), 2) 
 Next 
 StrToHexStr = Trim(strTemp) 
End Function 

Function Searchfordoclib(elElementName,cnvalue,XMLDoc)
Set xnSearchNodes = XMLDoc.selectNodes("//*[@" & elElementName & " = '" & cnvalue & "']")
If xnSearchNodes.length = 0 Then 
	Searchfordoclib = False
else
	Searchfordoclib = True
End if

End Function

sub Adddoclib(dn,hn,uf,uri,xnNodes)
	If Searchfordoclib("dn",dn,xdXmlDocument) = False then
		Set objnewEle = xdXmlDocument.createElement("docLib")
		objnewEle.setAttribute "uri",uri
		objnewEle.setAttribute "dn",dn
		objnewEle.setAttribute "hn", hn
		objnewEle.setAttribute "uf", uf
		xnNodes(0).appendchild objnewEle
		update = true
	Else
		wscript.echo "Dn exists " & DN
	End if
End sub
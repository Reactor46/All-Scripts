snServername = wscript.arguments(0)
mbMailboxName = wscript.arguments(1)

' Word Documents
ReDim wdattah(1)
wdattah(0) = "{DB13F464-2FAA-48F2-8D1B-ADB5ED4FD1F7}"
wdattah(1) = 22
'Excel Attachments
ReDim edattach(1)
edattach(0) = "{D549D2BB-E1BF-47DE-B713-784771F059A1}"
edattach(1) = 19
'PowerPoint Attachments
ReDim pptattach(1)
pptattach(0) = "{1E8ADCFB-AC2C-4FEF-ABB5-C5349A359CC8}"
pptattach(1) = 0
'PDF attachments
ReDim pdfattach(1)
pdfattach(0) = "{707D20D7-5EF8-47D7-B6C8-47FCB606EEB5}"
pdfattach(1) = 15
'Audio Attachments
ReDim sndattach(1)
sndattach(0) = "{B28E76F5-127B-4356-9150-D2A0B84E8DCE}"
sndattach(1) = 18
'Video
ReDim vdoattach(1)
vdoattach(0) = "{E633EC9C-9B29-4608-A4BA-CFBFA886702B}"
vdoattach(1) = 23
'Image Attachment
ReDim imgAttach(1)
imgAttach(0) = "{BB488D85-76FE-408F-9DD4-617041DBFDA6}"
imgAttach(1) = 13
'Zip Attachment
ReDim zipAttach(1)
zipAttach(0) = "{B4423425-54F1-304F-92F3-63451D3BFDB6}"
zipAttach(1) = 8

Set catDict = CreateObject("Scripting.Dictionary")
catDict.add "Word Attachment",wdattah
catDict.add "Excel Attachment",edattach
catDict.add "PowerPoint Attachment", pptattach
catDict.add "PDF Attachment", pdfattach
catDict.add "Audio Attachment", sndattach
catDict.add "Image Attachment", imgAttach
catDict.add "Video Attachment", vdoattach
catDict.add "Zip Attachment", zipAttach

set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
xdXmlDocument.async="false"
Set objSession   = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,snServername & vbLF & mbMailboxName
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
set cdocalendar = objSession.GetDefaultFolder(CdoDefaultFolderCalendar)
For Each soStorageItem in cdocalendar.HiddenMessages 
	If soStorageItem.Type = "IPM.Configuration.CategoryList" Then
		hexString = soStorageItem.fields(&h7C080102).Value
		xdXmlDocument.loadxml(hextotext(hexString))
		For Each cat In catDict
			catval = catDict(cat)
			If SearchforCategory("name",cat,xdXmlDocument) = True Or SearchforCategory("guid",catval(0),xdXmlDocument) Then
				wscript.echo "Category Name or GUID alread Exists " & cat
			Else
				wscript.echo "Adding category " & cat
				Call AddCategory(cat,catDict(cat),xdXmlDocument)
			End if
		next 
		nval = StrToHexStr(CStr(xdXmlDocument.xml))
		soStorageItem.fields(&h7C080102).Value = nval
		soStorageItem.update
	End if
Next 



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

Function SearchforCategory(elElementName,cnvalue,XMLDoc)
Set xnNodes = XMLDoc.selectNodes("//*[@" & elElementName & " = '" & cnvalue & "']")
If xnNodes.length = 0 Then 
	SearchforCategory = False
else
	SearchforCategory = True
End if

End Function

Function AddCategory(cnCategoryName,setarray,XMLDoc)
	Set xnNodes = XMLDoc.selectNodes("//categories")
	Set xnCatNodes = XMLDoc.selectNodes("//category")
	Set objnewCat = xnCatNodes(0).cloneNode(true)
	objnewCat.setAttribute "name",cnCategoryName
	objnewCat.setAttribute "guid",setarray(0)
	objnewCat.setAttribute "keyboardShortcut", 0
	objnewCat.setAttribute "color", setarray(1)
	objnewCat.setAttribute "usageCount", 0
	objnewCat.setAttribute "lastTimeUsedNotes","1601-01-01T00:00:00.000"
	objnewCat.setAttribute "lastTimeUsedJournal","1601-01-01T00:00:00.000"
	objnewCat.setAttribute "lastTimeUsedTasks","1601-01-01T00:00:00.000"
	objnewCat.setAttribute "lastTimeUsedContacts","1601-01-01T00:00:00.000"
	objnewCat.setAttribute "lastTimeUsedMail","1601-01-01T00:00:00.000"
	objnewCat.setAttribute "lastSessionUsed","1601-01-01T00:00:00.000"
	xnNodes(0).appendChild objnewCat
End Function
Const PR_FOLDER_WEBVIEWINFO = &H36DF0102

servername = "servername"
mailboxname = "mailbox"
flFolderName = "Webfolder"

homepage = "http://www.google.com.au/search?hl=en&q=coffee+tasting&btnG=Google+Search&meta=cr%3DcountryAU"
dwVersion = "02"
dwType = "00000001"
dwFlags = "00000001"
dwUnused = "00000000000000000000000000000000000000000000000000000000"
bData = AsciiToHex(homepage)
cbDataSize = cstr(ubound(bData)+1)
propval = dwVersion & dwType & dwFlags & dwUnused & "000000" & Hex(cbDataSize) & "000000" & Join(bData,"")



Set objSession   = CreateObject("MAPI.Session")

objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set objInbox = objSession.Inbox
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set CdoFolders = CdoFolderRoot.Folders

bFound = False
Set CdoFolder = CdoFolders.GetFirst
Do While (Not bFound) And Not (CdoFolder Is Nothing)
    If LCase(CdoFolder.Name) = Lcase(flFolderName) Then
       bFound = True
    Else
       Set CdoFolder = CdoFolders.GetNext
    End If
Loop

if bFound <> True then
	set webFld = CdoFolders.add(flFolderName)	
	webFld.fields.Add PR_FOLDER_WEBVIEWINFO, propval
	webFld.Update 
	wscript.echo "Folder Created"
else
	wscript.echo "Folder Exists Setting Homepage"
	Set ActionFolder = CdoFolder
	ActionFolder.fields.Add PR_FOLDER_WEBVIEWINFO, propval
	ActionFolder.Update 
end if 

Function AsciiToHex(sData)
 Dim i, aTmp()
 ReDim aTmp((Len(sData)*2) + 1)
 arnum = 0 
 For i = 1 To Len(sData)
  aTmp(arnum) = Hex(Asc(Mid(sData, i)))
  arnum = arnum + 1
  aTmp(arnum) = "00" 
  arnum = arnum + 1
 Next
  aTmp(arnum) = "00" 
  arnum = arnum + 1
  aTmp(arnum) = "00"
 ASCIItoHex = aTmp
End Function


PR_HAS_RULES = &H663A000B
PR_URL_NAME = &H6707001E
PR_CREATOR = &H3FF8001E
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\forwardingRules.csv",2,true)
wfile.writeline("Folder,FolderPath,Creator,AdressObject,SMTPForwdingAddress")
set objSession = CreateObject("MAPI.Session") 
strProfile = "mgnms01" & vbLf & "gscales" 
objSession.Logon "",,, False,, True, strProfile
Set objInfoStores = objSession.InfoStores 
set objInfoStore = objSession.GetInfoStore 
Set objpubstore = objSession.InfoStores("Public Folders")
Set objTopFolder = objpubstore.RootFolder 
call Loopfolders(objTopFolder,strfoldername) 
wfile.close


sub Loopfolders(cdofolder,strfoldername) 
set CdoFolders = cdofolder.Folders 
for each folder in CdoFolders
    call Loopfolders(folder,strfoldername)
    if folder.fields.item(PR_HAS_RULES) = true then
	Set objMessages = folder.HiddenMessages
	for Each objMessage in objMessages 
		if objMessage.type = "IPM.Rule.Message" then
 			call procrule(objMessage,folder.name,folder.fields.item(PR_URL_NAME).value)
		end if
	next
    end if 
next 
end sub

sub procrule(objmessage,foldername,folderpath)
frule = false
splitarry = split(hextotext(objmessage.fields.item(&H65EF0102)),chr(132),-1,1)
if ubound(splitarry) <> 0 then
	wscript.echo 
    	wscript.echo "Public Folder Name :" & foldername
	wscript.echo "Public Folder Path :" & folderpath
	wscript.echo "Rule Created By : " & objmessage.fields.item(PR_CREATOR).value
	fname = foldername
	fpath = folderpath
	creator = objmessage.fields.item(PR_CREATOR).value
	frule = true
end if
tfirst = 0
addcount = 1
for i = 0 to ubound(splitarry)
	addrrsplit = split(splitarry(i),chr(176),-1,1)
	for j = 0 to ubound(addrrsplit)
		addrcontsep = chr(3) & "0"
		if instr(addrrsplit(j),addrcontsep) then 
			if tfirst = 1 then addcount = addcount + 1
			wscript.echo 
			wscript.echo "Address Object :" & addcount
			redim Preserve resarray(1,1,1,1,1,addcount)
			resarray(1,0,0,0,0,addcount) = fname
			resarray(1,1,0,0,0,addcount) = fpath
			resarray(1,1,1,0,0,addcount) = creator		
			if instr(addrrsplit(j),"0/o=") then 
				resarray(1,1,1,1,0,addcount) = mid(addrrsplit(j),(instr(addrrsplit(j),"0/o=")+1),len(addrrsplit(j)))
				WScript.echo "ExchangeDN :" & mid(addrrsplit(j),(instr(addrrsplit(j),"0/o=")+1),len(addrrsplit(j)))
			else 
				WScript.echo "Address :" & mid(addrrsplit(j),3,len(addrrsplit(j)))
				resarray(1,1,1,1,0,addcount) = mid(addrrsplit(j),3,len(addrrsplit(j)))
			end if 
			tfirst = 1		
		end if
		smtpsep = Chr(254) & "9"
		if instr(addrrsplit(j),smtpsep) then 
			slen = instr(addrrsplit(j),smtpsep) + 2
			elen = instr(addrrsplit(j),chr(3))
			Wscript.echo "SMTP Forwarding Address : " & mid(addrrsplit(j),slen,(elen-slen))
			resarray(1,1,1,1,1,addcount) = mid(addrrsplit(j),slen,(elen-slen))
		end if
	next
next
if frule = true then
	for r = 1 to ubound(resarray,6)
		wfile.writeline(resarray(1,0,0,0,0,r) & "," & resarray(1,1,0,0,0,r) & "," & resarray(1,1,1,0,0,r) & "," & resarray(1,1,1,1,0,r) & "," & resarray(1,1,1,1,1,r))
	next
end if
end sub


Function hextotext(binprop)
arrnum = len(binprop)/2
redim aout(arrnum)
slen = 1
for i = 1 to arrnum
	if CLng("&H" & mid(binprop,slen,2)) <> 0 then
		aOut(i) = chr(CLng("&H" & mid(binprop,slen,2)))
		rem wscript.echo CLng("&H" & mid(binprop,slen,2)) & "," & chr(CLng("&H" & mid(binprop,slen,2)))
	end if
	slen = slen+2
next
hextotext = join(aOUt,"")
end function




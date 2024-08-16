Set doDictionaryObject = CreateObject("Scripting.Dictionary")
Set fso = CreateObject("Scripting.FileSystemObject")
Set fso = CreateObject("Scripting.FileSystemObject")
set RDOSession = CreateObject("Redemption.RDOSession")
tsize = 10
tnThreshold = 1800
servername = "servername"
mbMailbox = "mailboxname"
bfBaseFilename = "expMailbox"
pfFilePath = "c:\temp\"
fnFileName = ""
PST = ""
pstroot = ""
IPMRoot = ""
pfPstFile = ""
fNumber = 0
set wfile = fso.opentextfile(pfFilePath & bfBaseFilename & ".txt",2,true) 
RDOSession.LogonExchangeMailbox mbMailbox,servername
Set dfDeletedItemsFolder = RDOSession.GetDefaultFolder(3)
CreateNewPst()

wscript.echo fnFileName
wscript.echo "Enumerate Deleteditems"
for miLoop = 1 to IPMRoot.Folders.count
	ProcessItems(IPMRoot.Folders(miLoop))
	if IPMRoot.Folders(miLoop).Folders.count <> 0 then
		call Enumfolders(IPMRoot.Folders(miLoop),PstRootFolder,2)
	end if
next


function Enumfolders(FLDS,RootFolder,ltype)
for fl = 1 to FLDS.Folders.count
	if ltype = 1 then
		call ProcessFolderSub(FLDS.folders(fl),RootFolder)
	else
		ProcessItems(FLDS.folders(fl))
	end if
	wscript.echo FLDS.folders(fl).Name
	if FLDS.folders(fl).Folders.count <> 0 then
		if ltype = 1 then
			call Enumfolders(FLDS.folders(fl),FLDS.folders(fl).EntryID,1)
		else
			call Enumfolders(FLDS.folders(fl),FLDS.folders(fl).EntryID,2)
		end if
	end if
next
End function

Function CreateNewPst()
doDictionaryObject.RemoveAll
fNumber = fNumber + 1
fnFileName = pfFilePath & bfBaseFilename & "-" & fNumber & ".pst"
set PST = RDOSession.Stores.AddPSTStore(fnFileName, 1,  "Exported MailBox-" & now())
set pstroot = RDOSession.GetFolderFromID(PST.IPMRootFolder.EntryID,PST.EntryID)
For Each pstfld In PstRoot.folders
	If pstfld.Name = "Deleted Items" Then
		If fNumber = 1 Then
			doDictionaryObject.add dfDeletedItemsFolder.EntryID, pstfld.EntryID
			wscript.echo "Added Deleted Items Folder"
		End if
	End if
next
set IPMRoot = RDOSession.Stores.DefaultStore.IPMRootFolder
for fiLoop = 1 to IPMRoot.Folders.count
	if IPMRoot.Folders(fiLoop).Name <> "Deleted Items" then
		PstRootFolder = ProcessFolderRoot(IPMRoot.Folders(fiLoop),PST.IPMRootFolder.EntryID)
		if IPMRoot.Folders(fiLoop).Folders.count <> 0 then
			call Enumfolders(IPMRoot.Folders(fiLoop),IPMRoot.Folders(fiLoop).EntryID,1)
		end If
	Else
		if IPMRoot.Folders(fiLoop).Folders.count <> 0 then
			call Enumfolders(IPMRoot.Folders(fiLoop),IPMRoot.Folders(fiLoop).EntryID,1)
		end if
	end if
next
Set pfPstFile = fso.GetFile(fnFileName)

end function

function ProcessFolderRoot(Fld,parentfld)

set CDOPstfld = RDOSession.GetFolderFromID(parentfld,PST.EntryID)
wscript.echo fld.Name
Set newFolder = CDOPstfld.Folders.ADD(Fld.Name)	
ProcessFolder = newfolder.EntryID
newfolder.fields(&H3613001E) = Fld.fields(&H3613001E)

doDictionaryObject.add Fld.EntryID,newfolder.EntryID
end Function

function ProcessFolderSub(Fld,parentfld)

set CDOPstfld = RDOSession.GetFolderFromID(doDictionaryObject.item(parentfld),PST.EntryID)
wscript.echo fld.Name
Set newFolder = CDOPstfld.Folders.ADD(Fld.Name)	
ProcessFolder = newfolder.EntryID
newfolder.fields(&H3613001E) = Fld.fields(&H3613001E)

doDictionaryObject.add Fld.EntryID,newfolder.EntryID
end function

Sub ProcessItems(Fld)

	set dfDestinationFolder = RDOSession.GetFolderFromID(doDictionaryObject.item(Fld.EntryID),PST.EntryID)
	wscript.echo dfDestinationFolder.Name
	wfile.writeLine("Processing Folder : ") & dfDestinationFolder.Name
	for fiItemloop = 1 to Fld.Deleteditems.count
		on error resume next
		pfPredictednewSize = formatnumber((pfPstFile.size + Fld.Deleteditems(fiItemloop).size)/1048576,2,0,0,0)
		if err.number <> 0 Then
			Wscript.echo "Error Processing Item in " & Fld.Name
			wscript.echo  "EntryID of Item:" 
			wscript.echo  Fld.Deleteditems(fiItemloop).EntryID 
			wscript.echo  "Subect of Item:"
			wscript.echo   Fld.Deleteditems(fiItemloop).Subject
			Wfile.writeline("Error Processing Item in " & Fld.Name)
			Wfile.writeline("EntryID of Item:")
			Wfile.writeline(Fld.Deleteditems(fiItemloop).EntryID )
			Wfile.writeline("Subect of Item:")
			Wfile.writeline(Fld.Deleteditems(fiItemloop).Subject)
			err.clear
		end if
		If Int(pfPredictednewSize) >= Int(tsize) Then
			Wscript.echo "10 MB Exported"
			tsize = tsize + 10
		End if
		If  Int(pfPredictednewSize) >= Int(tnThreshold) Then
			wfile.writeLine("New PST about to be created - Destination - Number of Deleteditems: " & dfDestinationFolder.Deleteditems.count)
			CreateNewPst()
			set dfDestinationFolder = RDOSession.GetFolderFromID(doDictionaryObject.item(Fld.EntryID),PST.EntryID)
			call Fld.Deleteditems(fiItemloop).copyto(dfDestinationFolder)
			if err.number <> 0 then
				Wscript.echo "Error Processing Item in " & Fld.Name
				wscript.echo  "EntryID of Item:" 
				wscript.echo  Fld.Deleteditems(fiItemloop).EntryID 
				wscript.echo  "Subect of Item:"
				wscript.echo   Fld.Deleteditems(fiItemloop).Subject
				Wfile.writeline("Error Processing Item in " & Fld.Name)
				Wfile.writeline("EntryID of Item:")
				Wfile.writeline(Fld.Deleteditems(fiItemloop).EntryID )
				Wfile.writeline("Subect of Item:")
				Wfile.writeline(Fld.Deleteditems(fiItemloop).Subject)
				err.clear
			end if
		else
			call Fld.Deleteditems(fiItemloop).copyto(dfDestinationFolder)
			if err.number <> 0 then
				Wscript.echo "Error Processing Item in " & Fld.Name
				wscript.echo  "EntryID of Item:" 
				wscript.echo  Fld.Deleteditems(fiItemloop).EntryID 
				wscript.echo  "Subect of Item:"
				wscript.echo   Fld.Deleteditems(fiItemloop).Subject
				Wfile.writeline("Error Processing Item in " & Fld.Name)
				Wfile.writeline("EntryID of Item:")
				Wfile.writeline(Fld.Deleteditems(fiItemloop).EntryID )
				Wfile.writeline("Subect of Item:")
				Wfile.writeline(Fld.Deleteditems(fiItemloop).Subject)
				err.clear
			end if
		End if
		on error goto 0
	Next
	wfile.writeLine("Source - Number of Deleteditems : " & Fld.fields(&h36020003) & " Destination - Number of Deleteditems : " & dfDestinationFolder.messages.count)
end sub

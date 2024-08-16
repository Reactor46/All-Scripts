servername = "servername"
mailbox = "mailbox"
set CDOSession = CreateObject("MAPI.Session")
strProfile = servername & vbLf & mailbox
CDOSession.Logon "",,, False,, True, strProfile
set RDOSession = CreateObject("Redemption.RDOSession")
RDOSession.MAPIOBJECT = CDOSession.MAPIOBJECT
set PST = RDOSession.Stores.AddPSTStore("c:\temp\expExch.pst", 1, "Exported Deleted Items " & now())
set CDOPstRoot = CDOSession.getfolder(PST.IPMRootFolder.EntryID,PST.EntryID)
set IPMRoot = RDOSession.Stores.DefaultStore.IPMRootFolder
for i = 1 to IPMRoot.Folders.count
	pstid = ProcessFolder(IPMRoot.Folders(i),CDOPstRoot.ID)
	if IPMRoot.Folders(i).Folders.count <> 0 then
		call Enumfolders(IPMRoot.Folders(i),pstid)
	end if
next

function Enumfolders(FLDS,pstid)
for fl = 1 to FLDS.Folders.count
	subpstid = ProcessFolder(FLDS.folders(fl),pstid)
	if FLDS.folders(fl).Folders.count <> 0 then
		call Enumfolders(FLDS.folders(fl),subpstid)
	end if
next
End function

function ProcessFolder(Fld,parentfld)
if Fld.Name <> "Deleted Items" then
	set CDOPstfld = CDOSession.getfolder(parentfld,PST.EntryID)
	wscript.echo fld.Name
	Set newFolder = CDOPstfld.Folders.ADD(Fld.Name)	
	set RDOnewFolder = RDOSession.GetFolderFromID(newFolder.id,PST.EntryID)
	Wscript.echo newfolder.name
	for j = 1 to Fld.Deleteditems.count
		Fld.Deleteditems(j).copyto(RDOnewFolder)
		Wscript.echo "Copied Message " & Fld.Deleteditems(j).Subject
	next
	ProcessFolder = newfolder.id
	newfolder.Update 
else
	set CDOPstfld = CDOSession.getfolder(parentfld,PST.EntryID)
	wscript.echo fld.Name
	Set newFolder = CDOPstfld.Folders.ADD(Fld.Name & "-Mailbox")	
	set RDOnewFolder = RDOSession.GetFolderFromID(newFolder.id,PST.EntryID)
	Wscript.echo newfolder.name
	for j = 1 to Fld.Deleteditems.count
		Fld.Deleteditems(j).copyto(RDOnewFolder)
		Wscript.echo "Copied Message " & Fld.Deleteditems(j).Subject
	next
	ProcessFolder = newfolder.id
	newfolder.Update 
end if


End function



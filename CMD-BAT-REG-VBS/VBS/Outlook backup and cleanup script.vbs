'*********** Author : Somesh
'*********** Create date : 26-12-2013
'*********** Script Purpose: Outlook email backup and cleanup
'************ somesh_sn@hotmail.com

Option Explicit

dim strt, daysold
strt=now
Dim rootpath,path,foldr, strStoreName
Dim log, logline,strttime,endtime,status,foldcntr,closingnote,excludelist
Dim folder(75,5)
foldcntr=0

strStoreName=inputbox("Please enter the email address or display name as it appears in your Outlook","Email id")
if strStoreName="" then
msgbox "Invalid email, please re-run the script"
wscript.quit
End if

rootpath=inputbox("Please enter the path to save the emails","Path")
if rootpath="" Then
msgbox "Invalid path, please re-run the script"
wscript.quit
End if
Folder_exist(rootpath)
rootpath = rootpath & "\"

daysold= inputbox("Please enter the number of days","Older than days")
if daysold ="" then
msgbox "Invalid number, please re-run the script"
wscript.quit
elseif Not IsNumeric(daysold) then
msgbox "Please enter number only, please re-run the script"
Wscript.quit
End if
 
if msgbox("Do you want to exclude any folder(s) ?",36,"Confirmation")=vbyes then
	excludelist= Lcase(inputbox("To exclude any folder(s), please enter folder name as it appears in outlook separated by comma "","" ","Feed-in"))
End if

if msgbox("Are you sure you want to proceed ?",vbyesno,"confirmation")=vbno then
	Wscript.Quit
end if

msgbox "If you want to stop this process, please open Task Manager and kill ""wscript.exe"" under processes tab ",vbinformation,"Alert"


strttime=now
'Other declarations
Dim objOutlook,objNamespace,objStore,objRoot,objInbox,objSentItems
Dim objFSO,objHTAFile,objshell,objLOGFile,mailbody

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objStore = objNamespace.Stores.Item(strStoreName)
Set objRoot = objStore.GetRootFolder()
Set objInbox =objRoot.folders("Inbox")
Set objSentItems=objRoot.folders("Sent Items")

Dim objWorkingFolder,foldname
Dim colitems, olMsg, cnt
Dim objInputFile,size
olMsg = 3

'************ Call the function to save the email
Create_HTA_FILE
Set objshell=createobject("Wscript.Shell")
objShell.run ".\status1.hta"

Create_Log_File
mailbody="Outlook backup and clean-up tool has Started will send out an completion email, please do not run another instance"
SendEmail "Outlook backup and clean-up has Started : " & strttime,mailbody
Wscript.sleep "5000"
SaveEmail

Set objShell = CreateObject("WScript.Shell") 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objInputFile= objFSO.OpenTextFile("log.ini",1)
mailbody= objInputFile.readall
objInputFile.close

SendEmail "Outlook backup and clean-up has Finished : " & endtime,mailbody

Set objWorkingFolder=Nothing
Set objInbox =Nothing
Set objRoot = Nothing
Set objStore = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing
set objSentItems = Nothing

msgbox " emails are saved in location " & vbcr & rootpath,vbSystemModal,"Task Completed Successfully"
WScript.Quit

'**************** Script ends 

'############################ Functions ##########################

'*********************** Save email function 
Public Function SaveEmail()

path = rootpath & trim(objInbox) & "\"
folder_exist(path)
cnt=objInbox.items.count

Set colitems =objInbox.items
objWorkingFolder=objInbox.Name
foldname=objInbox.Name
SaveAndDeleteEmails() 
Set colitems =Nothing

Dim objSubFolder
For each objSubFolder in objInbox.folders
	path = rootpath & objInbox & "\" & trim(objSubFolder) & "\"
	folder_exist(path)
	cnt=objSubFolder.items.count
	Set colitems =objSubFolder.items
	objWorkingFolder=objInbox.Name & "\" & objSubFolder.Name
	foldname=LCase(objSubFolder.Name)
	SaveAndDeleteEmails() 
	Set colitems =Nothing
	if objSubFolder.folders.count > 1 then
		Dim fold
		For each fold in objSubFolder.folders
			path = rootpath & objInbox & "\" & trim(objSubFolder) & "\" & trim(fold) & "\"
			folder_exist(path)
			cnt=fold.items.count
			objWorkingFolder=objInbox.Name & "\" & objSubFolder.Name & "\" & fold.Name
			foldname=LCase(fold.Name)
			Set colitems =fold.items
			SaveAndDeleteEmails()
			Set colitems =Nothing
		NEXT
	END IF

Next 'folders loop

'*************** For Sent items
path = rootpath & trim(objSentItems) & "\"
folder_exist(path)
cnt=objSentItems.items.count
Set colitems = objSentItems.items
objWorkingFolder=objSentItems.name
foldname=LCase(objSentItems.name)
SaveAndDeleteEmails()
closingnote="<BR><B>Outlook backup and clean-up script has completed, you may now close this window</B><BR>"
status="<span style=""background-color: #90EE90"">Finished</span>"
endtime=now
Create_Log_File

Set objSubFolder=Nothing
Set fold =nothing

End Function

Sub SaveAndDeleteEmails()
foldcntr=foldcntr+1
folder(foldcntr-1,0)=objworkingFolder
folder(foldcntr-1,1)=cnt
folder(foldcntr-1,2)="-"
folder(foldcntr-1,3)="-"
folder(foldcntr-1,4)="Processing"
status = "<span style=""background-color: #FFFF00"">Running</span>"
Endtime ="Running"
Create_Log_File

If Instr(excludelist,foldname) <> 0 then
folder(foldcntr-1,2)="0"
folder(foldcntr-1,3)="0"
folder(foldcntr-1,4)="<span style=""background-color: #E6E6FA"">Excluded</span>"
Create_Log_File
Exit sub
End If

Dim counter
counter=0
if not cnt=0 then
Dim i
dim filename,tempfilename,fsize

for i=cnt to 0 step -1
	IF colitems(i).ReceivedTime < Dateadd("d",-daysold,Now) THEN
		filename =colitems(i).subject & " " & colitems(i).ReceivedTime & ".msg"
		tempfilename=CleanString(filename)
		fsize=fsize + colitems(i).size
		on error resume next
		colitems(i).SaveAs path & tempfilename,olMsg
		colitems(i).Delete
		counter =counter +1
			Else 
		EXIT FOR
	END IF
	folder(foldcntr-1,2)=counter
	folder(foldcntr-1,3)= Int((fsize/1024))
	Create_Log_File
next
end if
folder(foldcntr-1,2)=counter
folder(foldcntr-1,3)= Int((fsize/1024))
folder(foldcntr-1,4)="Finish"
size=size + fsize
Create_Log_File

End Sub

Function CleanString(strData)
    'Replace invalid strings.

    strData = Replace(strData, "´",   "'")
    strData = Replace(strData, "`",   "'")
    strData = Replace(strData, "{",   "(")
    strData = Replace(strData, "[",   "(")
    strData = Replace(strData, "]",   ")")
    strData = Replace(strData, "}",   ")")
    strData = Replace(strData, "  ",  " ")    'Replace two spaces with one space
    strData = Replace(strData, "   ", " ")    'Replace three spaces with one space    
    'Cut out invalid signs.
    strData = Replace(strData, ": ",  "_")    'Colan followded by a space
    strData = Replace(strData, ":",   "_")    'Colan with no space
    strData = Replace(strData, "/",   "_")
    strData = Replace(strData, "\",   "_")
    strData = Replace(strData, "*",   "_")
    strData = Replace(strData, "?",   "_")
    strData = Replace(strData, """",  "'")
    strData = Replace(strData, "<",   "_")
    strData = Replace(strData, ">",   "_")
    strData = Replace(strData, "|",   "_")
    CleanString = Trim(strData)
End Function

Function folder_exist(path)
on error resume next
Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

if NOT (objFSO.FolderExists(path)) then
objFSO.CreateFolder path
End if

END Function

Function Create_Log_File

Set objFSO=CreateObject("Scripting.FilesystemObject")
Set objLOGFile= objFSO.OpenTextFile(".\log.ini",2,true)
objLOGFile.writeline "<PRE style=""font-family:calibri;font-size:16px;"">Email account	: <B>" & strStoreName
objLOGFile.writeline "</B><BR>Start time		: " & strttime
objLOGFile.writeline "<BR>Status		: " & status
objLOGFile.writeline "<BR>End time		: " & endtime
objLOGFile.writeline "<BR>Path		: " & rootpath
objLOGFile.writeline "<BR>Number of days old emails to backup	: " & daysold
objLOGFile.writeline "<BR>Total Size freed up (Kb)	: " & (size/1024)
objLOGFile.writeline "</PRE><BR><Table border=""1""><style=""font-family:Times New Roman;""><TR><TD>Folder Name</TD><TD>Total emails </TD><TD>Processed</TD>" _  
			&"<TD>Size saved(Kb)</TD><TD>Status</TD></TR></TR>"
			
dim i
for i=0 to foldcntr -1
objLOGFile.writeline "<TR>"
objLOGFile.writeline "<TD>" & folder(i,0)
objLOGFile.writeline "</TD><TD>" & folder(i,1)
objLOGFile.writeline "</TD><TD>" & folder(i,2)
objLOGFile.writeline "</TD><TD>" & folder(i,3)
objLOGFile.writeline "</TD><TD>" & folder(i,4)
objLOGFile.writeline "</TR>"
Next
objLOGFile.writeline "</Table>"
objLOGFile.writeline closingnote
Set objFSO=Nothing
Set objLOGFILE= Nothing

End Function

Function Create_HTA_FILE
'on error resume next
Set objFSO=CreateObject("Scripting.FilesystemObject")
Set objHTAFile= objFSO.OpenTextFile(".\status1.hta",2,true)
objHTAFile.Writeline  "<html>"
objHTAFile.Writeline  "<head>"
objHTAFile.Writeline  "<H2>Status of the outlook emails backup and clean-up script</H2>"
objHTAFile.Writeline  "<title>Status - Auto Refreshed</title>"
objHTAFile.Writeline  "<HTA:APPLICATION "
objHTAFile.Writeline  "     ID=""objAutoRefresh"""
objHTAFile.Writeline  "	   APPLICATIONNAME=""Status - Auto Refreshed"""
objHTAFile.Writeline  "     SCROLL=""auto"""
objHTAFile.Writeline  "     SINGLEINSTANCE=""yes"""
objHTAFile.Writeline  ">"
objHTAFile.Writeline  "</head>"
objHTAFile.Writeline  "<SCRIPT LANGUAGE=""VBScript"">"
objHTAFile.Writeline  "	   Sub Window_OnLoad"
objHTAFile.Writeline  "		RefreshList "
objHTAFile.Writeline  "       iTimerID = window.setInterval(""RefreshList"", 1000)"
objHTAFile.Writeline  "	    End Sub"
objHTAFile.Writeline  "    Sub RefreshList"
objHTAFile.Writeline  "		strHTML="""""
objHTAFile.Writeline  "		Set objShell = CreateObject(""WScript.Shell"") "
objHTAFile.Writeline  "       	Set objFSO = CreateObject(""Scripting.FileSystemObject"")"
objHTAFile.Writeline  "		Set objInputFile= objFSO.OpenTextFile(""log.ini"",1)"
objHTAFile.Writeline  "		strHTML= objInputFile.readall"
objHTAFile.Writeline  "		objInputFile.close"
objHTAFile.Writeline  "       ProcessList.InnerHTML = strHTML"
objHTAFile.Writeline  "    End Sub"
objHTAFile.Writeline  "</SCRIPT>"
objHTAFile.Writeline  "<body><span id = ""ProcessList""></span>"
objHTAFile.Writeline  "</body>"
objHTAFile.Writeline  "<sub>"
objHTAFile.Writeline  "-- <BR>"
objHTAFile.Writeline  "Scripted by Somesh</sub>"
objHTAFile.Writeline  "</html>"

objHTAFile.close

Set objFSO=Nothing
Set objHTAFile=Nothing

END Function 'Somesh

Function SendEmail(subject,mailbody)
Dim objemail
on error resume next
Set objemail=objOutlook.createitem(0)
objemail.to=strStoreName
objemail.SentOnBehalfOfName=strStoreName

objemail.subject=subject
objemail.htmlbody=mailbody
objemail.send

Set objemail = nothing

End Function

'********************* END ***************************
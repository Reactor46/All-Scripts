set shell = createobject("wscript.shell")
domainname = "domain.com"
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
set conn1 = createobject("ADODB.Connection")
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\" & wscript.arguments(0) & ".csv"
set wfile = fso.opentextfile(fname,2,true)
     wfile.writeline("File Group,File Type,File Extension,File Name, File Size, Email Location,Message Subject,Sent From,Date Sent")
public exlist
public maxsize
public datefrom
public dateto
maxsize = wscript.arguments(1)
datefrom = wscript.arguments(2) & "T00:00:00Z"
dateto = wscript.arguments(3) & "T00:00:00Z"
exlist = ","
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
set objgrandchild = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS RFileType, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS CFileext, " & _
           "  	  NEW adVarChar(255) AS CFileType, " & _
           "  	  NEW adVarChar(255) AS CFileDesc, " & _
           " ((SHAPE APPEND  " & _
	   "      NEW adVarChar(255) AS GCFileext, " & _
	   "      NEW adVarChar(255) AS GCFilename, " & _
	   "      NEW adVarChar(255) AS GCFilesize, " & _
	   "      NEW adVarChar(255) AS GCMessLocation, " & _
	   "      NEW adVarChar(255) AS GCMessSubject, " & _
	   "      NEW adVarChar(255) AS GCMessFrom, " & _
	   "      NEW adVarChar(255) AS GCDateSent) " & _
	   "   RELATE CFileext TO GCFileext) AS MOWMI" & _
	   ")" & _
           "      RELATE RFileType TO CFileType) AS rsSOMO " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
rem *********** Add Document Types
call adddoctype(objParentRS,"Microsoft Office Documents")
call adddoctype(objParentRS,"Compressed Files")
call adddoctype(objParentRS,"Acorbat Files")	
call adddoctype(objParentRS,"Executables and Installers")	
call adddoctype(objParentRS,"Sound Files")
call adddoctype(objParentRS,"Video Files")
call adddoctype(objParentRS,"Image Files")
call adddoctype(objParentRS,"Attached Email Message")
call adddoctype(objParentRS,"Unclassified Files")
rem ********************************************
rem *********** Add Document Extensions"
Set objChildRS = objParentRS("rsSOMO").Value
call addfileexts(objChildRS)
Set objgrandchild = objChildRS("MOWMI").Value
sConnString = "file://./backofficestorage/" & domainname
sConnString = sConnString & "/mbx/" & wscript.arguments(0) & "/"
WScript.Echo sConnString
call RecurseFolder(sConnString,objChildRS,objgrandchild)

objParentRS.MoveFirst
Do While Not objParentRS.EOF
     wscript.echo objParentRS(0)
     Set objChildRS = objParentRS("rsSOMO").Value 
     Do While Not objChildRS.EOF 
	Set objgrandchild = objChildRS("MOWMI").Value
	if objgrandchild.recordcount <> 0 then 
		wscript.echo "	" & objChildRS(1) & "  " & objChildRS(2)
		wfile.writeline(objChildRS(1) & "," & objChildRS(2))
	end if
	Do While Not objgrandchild.EOF 
		wscript.echo "		" & objgrandchild(1) & "  " & objgrandchild(2) & "  " & objgrandchild(3) & "  "  & objgrandchild(4) & "  "
		wfile.writeline(",," & objgrandchild(0) & "," & objgrandchild(1) & "," & objgrandchild(2) & "," & objgrandchild(3) & ","  & objgrandchild(4) & "," & objgrandchild(5) & "," & objgrandchild(6))
		objgrandchild.movenext
	loop
	objChildRS.movenext
     loop
     objParentRS.MoveNext
Loop

Public Sub RecurseFolder(sConnString,objChildRS,objgrandchild)
  
   sSQL = "SELECT ""http://schemas.microsoft.com/mapi/proptag/x0e080003"", "
   sSQL = sSQL & """DAV:href"", "
   sSQL = sSQL & """DAV:hassubs"" "
   sSQL = sSQL & "FROM SCOPE ('SHALLOW TRAVERSAL OF """ & sConnString
   sSQL = sSQL & """') WHERE ""DAV:isfolder"" = true"
   Set oConn = CreateObject("ADODB.Connection")
   oConn.Provider = "Exoledb.DataSource"
   oConn.Open sConnString
   Set oRecSet = CreateObject("ADODB.Recordset")
   oRecSet.CursorLocation = 3
   oRecSet.Open sSQL, oConn.ConnectionString
   oRecSet.MoveFirst
   While oRecSet.EOF <> True
      call  procfolder(oRecSet.Fields("DAV:HREF") & "/",objgrandchild,objChildRS)
      If oRecSet.Fields.Item("DAV:hassubs") = True then
         call RecurseFolder(oRecSet.Fields.Item("DAV:href"),objChildRS,objgrandchild)
      End If
      oRecSet.MoveNext
   wend
   oRecSet.Close
   oConn.Close
   Set oRecSet = Nothing
   Set oConn = Nothing
End Sub


sub procfolder(folderurl,objgrandchild,objChildRS)
rem ***********Mailbox Section
WScript.Echo folderurl
Set Rs = CreateObject("ADODB.Recordset")
set Rec = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open folderurl, ,3
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & folderurl & """') " 
SSql = SSql & " Where ""DAV:isfolder"" = false AND ""urn:schemas:httpmail:hasattachment"" = true AND ""DAV:ishidden"" = false "           
SSql = SSql & "AND ""urn:schemas:httpmail:datereceived"" > CAST(""" & datefrom & """ as 'dateTime') " _
& "AND ""urn:schemas:httpmail:datereceived"" < CAST(""" & dateto & """ as 'dateTime')"  
Rs.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
i = 1
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
	call procmail(Rs.Fields("DAV:href").Value,objgrandchild,objChildRS)
	rs.movenext
wend
end if
rs.close
end sub

sub procmail(murl,objgrandchild,objChildRS)
rem on error resume next
set msg = createobject("cdo.message") 
msg.datasource.open murl
lcMaloc = mid(msg.fields("Dav:parentname").value,instr(msg.fields("Dav:parentname").value,"/mbx/")+5,len(msg.fields("Dav:parentname").value))
stSenttime = msg.fields("urn:schemas:httpmail:datereceived").value
fnFromName =  msg.fields("urn:schemas:httpmail:fromname").value
feFromEmail =  replace(replace(msg.fields("urn:schemas:httpmail:fromemail").value,"<",""),">","")
toToEmail = msg.fields("urn:schemas:mailheader:to").value
sjSubject = msg.Subject
i = 1
set objattachments = msg.attachments 
for each objattachment in objattachments 
if objAttachment.ContentMediaType = "message/rfc822" then
	set msg1 = createobject("cdo.message") 
	msg1.datasource.OpenObject objattachment, "ibodypart"
	fnFileName = msg1.subject & "(" & i & ")" & ".eml"
	set stm = msg1.getstream
else 
	fnFileName = objattachment.filename
	set stm = objAttachment.GetDecodedContentStream 
end if 
if fnFileName <> "" and clng(stm.size/1024) > clng(maxsize) then
	wscript.echo fnFileName
	fatt1 = len(fnFileName)
	fatt2 = fatt1 - 2
	rtime = formatdatetime(stSenttime)
	attname = UCASE(fnFileName)
	objgrandchild.addnew 
	objgrandchild("GCFileext") = mid(attname,fatt2,3)
	objgrandchild("GCFilename") = right(replace(fnFileName,",",""),255)
	objgrandchild("GCFilesize") = right(replace(formatnumber(stm.size/1024,2),",",""),255)
	objgrandchild("GCMessLocation") = right(replace(lcMaloc,",",""),255)
	objgrandchild("GCMessSubject") = right(replace(sjSubject,",",""),255)
	objgrandchild("GCMessFrom") = feFromEmail
	objgrandchild("GCDateSent") = rtime
	objgrandchild.update
	elistchk = "," & mid(attname,fatt2,3) & ","
	if instr(exlist,elistchk) = 0 then
		call adddocext(objChildRS,"Unclassified Files",mid(attname,fatt2,3),"Unknown")	
	end if 
end if
i = i + 1
next 
set msg = nothing 


end sub

sub adddoctype(objParentRS,doctype)

objParentRS.addnew 
objParentRS("RFileType") = doctype
objParentRS.update

end sub


sub adddocext(objChildRS,doctype,ext,fdesc)

objChildRS.addnew 
objChildRS("CFileType") = doctype
objChildRS("CFileext") = ext
objChildRS("CFileDesc") = fdesc
exlist = exlist & "," & ext & ","
objChildRS.update

end sub 

sub addfileexts(objChildRS)
call adddocext(objChildRS,"Microsoft Office Documents","DOC","Microsoft Word Document")
call adddocext(objChildRS,"Microsoft Office Documents","DOT","Microsoft Word Template")
call adddocext(objChildRS,"Microsoft Office Documents","XLS","Microsoft Excel Spreedsheet")
call adddocext(objChildRS,"Microsoft Office Documents","PPT","Microsoft Powerpoint Presentation")
call adddocext(objChildRS,"Microsoft Office Documents","PPS","Microsoft Powerpoint Slide Show")
call adddocext(objChildRS,"Microsoft Office Documents","MDB","Microsoft Access Database")
call adddocext(objChildRS,"Microsoft Office Documents","ADP","Microsoft Access Project")
call adddocext(objChildRS,"Microsoft Office Documents","VSD","Microsoft Visio Diagram")
call adddocext(objChildRS,"Microsoft Office Documents","ONE","Microsoft OneNote Note")
call adddocext(objChildRS,"Microsoft Office Documents","RTF","Rich Text Format file")
call adddocext(objChildRS,"Microsoft Office Documents","TXT","Text File")
call adddocext(objChildRS,"Microsoft Office Documents","CSV","Comma seperated File")
call adddocext(objChildRS,"Compressed Files","ZIP","Zip Compressed  File")
call adddocext(objChildRS,"Compressed Files","TAR","nix Tar Compressed File")
call adddocext(objChildRS,"Compressed Files","ARG","Arg Compressed  File")
call adddocext(objChildRS,"Compressed Files","RAR","RAR Compressed File")
call adddocext(objChildRS,"Compressed Files","ACE","ACE Compressed File")
call adddocext(objChildRS,"Compressed Files","BHX","Binary Hex Compressed File")
call adddocext(objChildRS,"Acorbat Files","PDF","Adobe Acrobat File")
call adddocext(objChildRS,"Executables and Installers","EXE","Executable File")
call adddocext(objChildRS,"Executables and Installers","MSI","Windows Installer")
call adddocext(objChildRS,"Sound Files","WAV","Wave File")
call adddocext(objChildRS,"Sound Files","MP3","MPeg3 Sound file")
call adddocext(objChildRS,"Sound Files","WMA","Windows Media file")
call adddocext(objChildRS,"Sound Files","WMV","Windows Media file")
call adddocext(objChildRS,"Sound Files","SND","Windows Sound File")
call adddocext(objChildRS,"Sound Files",".AU","AU Sound File")
call adddocext(objChildRS,"Sound Files","RPM","Real Audio Sound File")
call adddocext(objChildRS,"Sound Files","MID","MIDI Audio Sound File")
call adddocext(objChildRS,"Sound Files",".RM","Real Audio Sound File")
call adddocext(objChildRS,"Sound Files",".RA","Real Audio Sound File")
call adddocext(objChildRS,"Sound Files","ASF","Advanced Streaming format File")
call adddocext(objChildRS,"Video Files","AVI","AVI Video format file")
call adddocext(objChildRS,"Video Files","MPG","MPG Video format file")
call adddocext(objChildRS,"Video Files","MOV","MOV Video format file")
call adddocext(objChildRS,"Video Files","IVX","DIVX Video format file")
call adddocext(objChildRS,"Video Files","PG4","MPG4 Video format file")
call adddocext(objChildRS,"Video Files","SWF","Shockwave format file")
call adddocext(objChildRS,"Image Files","JPG","JPG picture file")
call adddocext(objChildRS,"Image Files","BMP","Bit Map picture file")
call adddocext(objChildRS,"Image Files","GIF","Gif picture file")
call adddocext(objChildRS,"Image Files","PNG","Portable Network graphics picture file")
call adddocext(objChildRS,"Image Files","TIF","Tag Image picture file")
call adddocext(objChildRS,"Image Files","IFF","Tag Image picture file")
call adddocext(objChildRS,"Image Files","WMF","Windows Metafile file")
call adddocext(objChildRS,"Image Files","EMF","Enhanced Metafile file")
call adddocext(objChildRS,"Image Files","PEG","Enhanced Metafile file")
call adddocext(objChildRS,"Attached Email Message","EML","Attached Email Message")
call adddocext(objChildRS,"Attached Email Message","ICS","Attached Calendar file")


end sub

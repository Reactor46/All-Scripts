on error resume next
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
set conn1 = createobject("ADODB.Connection")
set req = createobject("microsoft.xmlhttp")
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\" & wscript.arguments(1) & ".csv"
set wfile = fso.opentextfile(fname,2,true)
     wfile.writeline("File Group,File Type,File Extension,File Name, File Size, Email Location,Message Subject,Sent From,Date Sent")
public exlist
public maxsize
public datefrom
public dateto
maxsize = wscript.arguments(2)
datefrom = wscript.arguments(3) & "T00:00:00Z"
dateto = wscript.arguments(4) & "T00:00:00Z"
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
sConnString = "https://" & wscript.arguments(0) & "/exchange/" & wscript.arguments(1) & "/NON_IPM_SUBTREE"
Set reqfba = CreateObject("Microsoft.xmlhttp")
domain = "domain"
strpassword = "password"
strusername =  domain & "\" & "username"
szXml = "destination=https://" & wscript.arguments(0) & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
reqfba.Open "post", "https://" & wscript.arguments(0) & "/exchweb/bin/auth/owaauth.dll", False
reqfba.send szXml
reqhedrarry = split(reqfba.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
call RecurseFolder(sConnString,objChildRS,objgrandchild)

objParentRS.MoveFirst
Do While Not objParentRS.EOF
     wscript.echo objParentRS(0)
     Set objChildRS = objParentRS("rsSOMO").Value 
     Do While Not objChildRS.EOF 
	attsum = 0
        attnum = 0
	Set objgrandchild = objChildRS("MOWMI").Value
	attnum = objgrandchild.recordcount
	if objgrandchild.recordcount <> 0 then 
		wscript.echo "	" & objChildRS(1) & "  " & objChildRS(2)
		wfile.writeline(objChildRS(1) & "," & objChildRS(2))
	end if
	Do While Not objgrandchild.EOF 
		wscript.echo "		" & objgrandchild(1) & "  " & objgrandchild(2) & "  " & objgrandchild(3) & "  "  & objgrandchild(4) & "  "
		wfile.writeline(",," & objgrandchild(0) & "," & objgrandchild(1) & "," & objgrandchild(2) & "," & objgrandchild(3) & ","  & objgrandchild(4) & "," & objgrandchild(5) & "," & objgrandchild(6))
		attsum = attsum + clng(objgrandchild("GCFilesize"))
		objgrandchild.movenext
	loop
	if attnum <> 0 then 
		attachmentsumy = attachmentsumy & objChildRS(1) & "," & objChildRS(2) & "," & objChildRS(0) & "," & attnum & "," & attsum & vbcrlf
	end if
	objChildRS.movenext
     loop
     objParentRS.MoveNext
Loop
wscript.echo 
wscript.echo "Attachment Summary"
wscript.echo
wscript.echo attachmentsumy
wfile.writeline
wfile.writeline "Attachment Summary"
wfile.writeline
wfile.writeline attachmentsumy

Public Sub RecurseFolder(sUrl,objChildRS,objgrandchild)
  
   Set oXMLHttp = CreateObject("Microsoft.xmlhttp")
   oXMLHttp.open "SEARCH", sUrl, False, "", ""
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and ""DAV:ishidden"" = false and ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 1"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"
   oXMLHttp.setRequestHeader "Content-Type", "text/xml"
   oXMLHttp.setRequestHeader "Translate", "f"
   oXMLHttp.setRequestHeader "Depth", "0"
   oXMLHttp.SetRequestHeader "cookie", reqsessionID
   oXMLHttp.SetRequestHeader "cookie", reqCadata
   oXMLHttp.setRequestHeader "Content-Length", "" & Len(sQuery)
   oXMLHttp.send sQuery
   Set oXMLDoc = oXMLHttp.responseXML
   Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
   For i = 0 to oXMLSizeNodes.length - 1
      call procfolder(oXMLHREFNodes.Item(i).nodeTypedValue,objgrandchild,objChildRS)
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         call RecurseFolder(oXMLHREFNodes.Item(i).nodeTypedValue,objChildRS,objgrandchild)
      End If
   Next
End Sub


sub procfolder(strURL,objgrandchild,objChildRS)

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """urn:schemas:httpmail:datereceived"", ""urn:schemas:httpmail:fromname"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """urn:schemas:httpmail:hasattachment"" = True AND " 
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &lt; CAST(""" & dateto & """ as 'dateTime') AND "
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &gt; CAST(""" & datefrom & """ as 'dateTime')</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   set oSubject = oResponseDoc.getElementsByTagName("d:subject")
   set odatereceived = oResponseDoc.getElementsByTagName("d:datereceived")
   set ofromemail = oResponseDoc.getElementsByTagName("d:fromemail")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	set oNode2 = oSubject.nextNode
	set oNode3 = odatereceived.nextNode
	set oNode4 = ofromemail.nextNode
	call embedattach(oNode1.Text,oNode2.Text,oNode3.Text,oNode4.Text,oNode.Text,objgrandchild,objChildRS)
   Next	
Else
End If

end sub

sub embedattach(objhref,subject,daterecieved,recievedfrom,davdisplay,objgrandchild,objChildRS)
req.open "X-MS-ENUMATTS", objhref, false, "", ""
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.send
If req.status > 207 Or req.status < 207 Then
Else
    set resDoc1 = req.responseXML
    Set objHrefNodeList = resDoc1.getElementsByTagName("a:href")
    Set objattachmethod = resDoc1.getElementsByTagName("d:x37050003")
    Set objcnval = resDoc1.getElementsByTagName("f:cn")
    set objattachname = resDoc1.getElementsByTagName("d:x3704001f")
    set objattachfilename = resDoc1.getElementsByTagName("e:attachmentfilename")
    set objattsize = resDoc1.getElementsByTagName("d:x0e200003")
    If objHrefNodeList.length > 0 Then
    For f = 0 To (objHrefNodeList.length -1)
        set objHrefNode1 = objHrefNodeList.nextNode
       	set objNodef = objattachmethod.nextnode
	if objattachmethod.length <> 0 then
       	if objNodef.Text = 5 then
            call embedattach(objHrefNode1.Text,subject,daterecieved,recievedfrom,davdisplay,objgrandchild,objChildRS)
       	else
	    settodav = 0
            set objNode1f = objattachfilename.nextNode
	    if objattachfilename.length = 0 then
		if objattachname.length = 0 then
			set objNode1f = objcnval.nextNode
			if objcnval.length = 0 then settodav = 1
		else
	    		set objNode1f = objattachname.nextNode
		end if
	    end if
	    fnFileName = objNode1f.Text
	    if err.number <> 0 then wscript.echo "error" & objHrefNode1.Text
	    set objNode3f = objattsize.nextnode
	    attsize = objNode3f.Text

	    if fnFileName <> "" and clng(attsize/1024) > clng(maxsize) then
		wscript.echo fnFileName
		fatt1 = len(fnFileName)
		lcMaloc = replace(objhref,"%E2%80%99S%20"," ")
		lcMaloc = replace(unescape(lcMaloc),davdisplay,"")
		lcMaloc = replace(lcMaloc,"http://" & wscript.arguments(0) & "/exchange/" & wscript.arguments(1) & "/","")
		wscript.echo lcMaloc
		fatt2 = fatt1 - 2
		attname = UCASE(fnFileName)
		rtime = dateserial(mid(daterecieved,1,4),mid(daterecieved,6,2),mid(daterecieved,9,2)) & " " & mid(daterecieved,12,8)
		objgrandchild.addnew 
		objgrandchild("GCFileext") = mid(attname,fatt2,3)
		objgrandchild("GCFilename") = right(replace(fnFileName,",",""),254)
		objgrandchild("GCFilesize") = replace(formatnumber(attsize/1024,2),",","")
		objgrandchild("GCMessLocation") = right(replace(lcMaloc,",",""),254)
		objgrandchild("GCMessSubject") = right(replace(subject,",",""),254)
		objgrandchild("GCMessFrom") = recievedfrom
		objgrandchild("GCDateSent") = dateadd("h",toffset,formatdatetime(rtime,0))
		objgrandchild.update
		elistchk = "," & mid(attname,fatt2,3) & ","
		if instr(exlist,elistchk) = 0 then
			call adddocext(objChildRS,"Unclassified Files",mid(attname,fatt2,3),"Unknown")	
		end if
	end if
	end if 
       end if
    next
Else
End If
End If
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

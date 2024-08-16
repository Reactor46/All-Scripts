FeedURL = wscript.arguments(0)
PubFolderURL = wscript.arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
set logfile = fso.opentextfile("c:\temp\RssReadFeedLog.txt",8,true)


logfile.writeline "*****Version 1.4*****"

set Conn = createobject("ADODB.Connection")
conn.provider = "ExOLEDB.Datasource"
conn.Open PubFolderURL, "", "", -1 

LastPost = GetLastPost(PubFolderURL)

if LastPost = "" then
   logfile.writeline "Full Sync"
   SyncType = "Full"
   ModSince = "0"
   Etag = "0"
else
   logfile.writeline "Partial Sync"
   SyncType = "Part"
   LastPostarry = split(LastPost,";")
   ModSince = LastPostarry(0)
   Etag = LastPostarry(1)
end if

Set XMLreq = CreateObject("Msxml2.XMLHTTP.4.0")
XMLreq.open "GET", FeedURL, true
XMLreq.setRequestHeader "Accept-Encoding", "gzip, deflate"
XMLreq.setRequestHeader "If-Modified-Since", ModSince
XMLreq.setRequestHeader "If-None-Match", Etag
call XMLreq.send
lexit = 0
do until lexit = 30
	If XMLreq.readyState <> 4 then
    		wscript.sleep 1000
    		lexit = lexit + 1
	else
   		lexit = 30
	End If
loop
If XMLreq.readyState <> 4 then logfile.writeline "Request Timeout Occured after 30 Seconds"
logfile.writeline XMLreq.status
if  XMLreq.status = "200" then
	if instr(XMLreq.getallresponseheaders,"Content-Encoding") then logfile.writeline "Compressed Stream: " & XMLreq.getResponseHeader("Content-Encoding")
	if instr(XMLreq.getallresponseheaders,"Last-Modified") then
		logfile.writeline XMLreq.getResponseHeader("Last-Modified")
		if instr(XMLreq.getallresponseheaders,"ETag") then
			logfile.writeline XMLreq.getResponseHeader("ETag")
			LastPostValue = XMLreq.getResponseHeader("Last-Modified") & ";" & XMLreq.getResponseHeader("ETag")
		else
			LastPostValue = XMLreq.getResponseHeader("Last-Modified") & ";0"
		end if
	else
		LastPostValue = "0;0"
		logfile.writeline "Doesn't Support Condition Get"
	end if
	call SetLastPost(PublicFolderURL,LastPostValue)
	set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
	xdXmlDocument.async = false
	xdXmlDocument.validateOnParse = false 
	xdXmlDocument.load(XMLreq.ResponseBody) 
	if Modsince <> "0" then 
		LastPostnum = numdateit(cdate(replace(replace(mid(ModSince,instr(ModSince,",")+2,len(ModSince))," GMT",""),"UT","")))
	end if
	frun = 0
	Set xnFeedNodes = xdXmlDocument.selectNodes("feed")
	if xnFeedNodes.length = 1 then
		call AtomFeed(SyncType,xdXmlDocument,PubFolderURL,LastPostnum)
	else
		Set xnRssNodes = xdXmlDocument.selectNodes("rss")
		if xnRssNodes.length <> 0 then
			call RssFeed(SyncType,xdXmlDocument,PubFolderURL,LastPostnum)
		else
			call RDFFeed(SyncType,xdXmlDocument,PubFolderURL,LastPostnum)
		end if
	end if
	
else
	logfile.writeline "Feed Not UPdated"
	logfile.writeline "http request status " & XMLreq.status 
end if


Function GetLastPost(PublicFolderURL)

Set rec = Createobject("ADODB.Record")
rec.open cstr(PublicFolderURL),conn,3
logfile.writeline FeedURL
GetLastPost = rec.fields(cstr(escape(FeedURL))).Value
logfile.writeline "Public Folder Property Value: " & GetLastPost
rec.close

end function

Sub SetLastPost(PublicFolderURL,PubDate)
if PubDate = "" then PubDate = numdateit(now())
Set rec = Createobject("ADODB.Record")
rec.open cstr(PublicFolderURL),conn,3
rec.fields(cstr(escape(FeedURL))) = PubDate
rec.fields.update
rec.close

end Sub

function numdateit(datetocon)
	strDateTime = year(datetocon) 
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) 
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & formatdatetime(datetocon,4) & "00"
	numdateit = replace(strDateTime,":","")
end function

function anumit(datetocon)
	datetocon = replace(replace(replace(replace(datetocon,"-",""),":",""),"T",""),"Z","")
	anumit = datetocon
end function

Sub AtomFeed(SyncType,xdXmlDocument,PublicFolderURL,LastPost)
	dim modified,title,hreflink,content,author
	if SyncType = "Full" then
		Set xnEntryNodes = xdXmlDocument.selectNodes("feed/entry")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename
					case "author" author = xnEntryDet.text
					case "id" id = xnEntryDet.text
					case "issued" issued =  xnEntryDet.text
					case "modified" modified = xnEntryDet.text
					case "updated" modified = xnEntryDet.text
					case "created" created =  xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    if  instr(xnEntryDet.xml,"type=""html""") then
								content = xnEntryDet.text
							    else
								content = xnEntryDet.xml
							    end if	
						       end if
					case "summary" if instr(xnEntryDet.xml,"mode=""escaped""") then
							    if content = "" then content = xnEntryDet.text
						       else
							    if  instr(xnEntryDet.xml,"type=""html""") then
								if content = "" then content = xnEntryDet.text
							    else
								if content = "" then content = xnEntryDet.xml
							    end if	
						       end if
					case "title" title = xnEntryDet.text
					case "link" if xnEntryDet.attributes.getNamedItem("rel").nodeValue = "alternate" then
						   	 hreflink = xnEntryDet.attributes.getNamedItem("href").nodeValue
						   	 if instr(xnEntryDet.xml,"title=""") then
						   	 	title = xnEntryDet.attributes.getNamedItem("title").nodeValue
						  	 end if
						    end if
				end select
			next
			call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
		next
	else 
		Set xnEntryNodes = xdXmlDocument.selectNodes("feed/entry")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename
					case "author" author = xnEntryDet.text
					case "id" id = xnEntryDet.text
					case "issued" issued =  xnEntryDet.text
					case "modified" modified = xnEntryDet.text
					case "updated" modified = xnEntryDet.text
					case "created" created =  xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    if  instr(xnEntryDet.xml,"type=""html""") then
								content = xnEntryDet.text
							    else
								content = xnEntryDet.xml
							    end if	
							end if
					case "summary" if instr(xnEntryDet.xml,"mode=""escaped""") then
							    if content = "" then content = xnEntryDet.text
						       else
							    if  instr(xnEntryDet.xml,"type=""html""") then
								if content = "" then content = xnEntryDet.text
							    else
								if content = "" then content = xnEntryDet.xml
							    end if	
						       end if
					case "title" title = xnEntryDet.text
					case "link"  if xnEntryDet.attributes.getNamedItem("rel").nodeValue = "alternate" then
						   	 hreflink = xnEntryDet.attributes.getNamedItem("href").nodeValue
						   	 if instr(xnEntryDet.xml,"title=""") then
						   	 	title = xnEntryDet.attributes.getNamedItem("title").nodeValue
						  	 end if
						     end if
				end select
			next
			if anumit(created) > lastpost then
				call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
				logfile.writeline "post added " & created & "  " & lastpost
			else
				if modified <> "" then
					if anumit(modified)  > lastpost then
						call ModifyPost(PublicFolderURL,modified,title,hreflink,content,author,id)
						logfile.writeline "post added Modified " & modified & "  " & lastpost
					end if 
				else
					logfile.writeline "post not added " & created & "  " & lastpost
				end if
				
			end if
		next
	
	end if
	logfile.writeline "Atom Feed"
end Sub

Sub RssFeed(SyncType,xdXmlDocument,PublicFolderURL,LastPost)
	dim modified,title,hreflink,content,author
	if SyncType = "Full" then
		Set xnEntryNodes = xdXmlDocument.selectNodes("rss/channel/item")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename
					case "author" author = xnEntryDet.text
					case "dc:creator" author = xnEntryDet.text
					case "guid" id = xnEntryDet.text
					case "dc:date" created =  xnEntryDet.text
					case "pubDate" created =  xnEntryDet.text
					case "description" content = xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    content = xnEntryDet.xml	
						       end if
					case "dc:title" title = xnEntryDet.text
					case "title" title = xnEntryDet.text
					case "link" hreflink = xnEntryDet.text
				end select
			next
			call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
		next
	else 
		Set xnEntryNodes = xdXmlDocument.selectNodes("rss/channel/item")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename				
					case "author" author = xnEntryDet.text
					case "dc:creator" author = xnEntryDet.text
					case "guid" id = xnEntryDet.text
					case "dc:date" created =  xnEntryDet.text
					case "pubDate" created = xnEntryDet.text
					case "description" content = xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    content = xnEntryDet.xml	
						       end if
					case "dc:title" title = xnEntryDet.text
					case "title" title = xnEntryDet.text
					case "link" hreflink = xnEntryDet.text
				end select
			next
			if created <> "" then
				c1 = mid(created,instr(created,",")+2,len(created))
				c2 = replace(c1,"GMT","")
				c3 = replace(c2,"UT","")
				c4 = split(c3)
				c5 = c4(0) & " " & c4(1) & " " & c4(2) & " " & c4(3)
				createdcvn = numdateit(cdate(c5))
				logfile.writeline "RSSFEED CVN = " & createdcvn
			end if
			if createdcvn > lastpost then
				call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
				logfile.writeline "post added " & created & "  " & lastpost
			else
				if created = "" then 
					call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
				else
					logfile.writeline "post not added " & created & "  " & lastpost
				end if
			end if
		next
	
	end if
	logfile.writeline "RSS Feed"
end Sub

Sub RDFFeed(SyncType,xdXmlDocument,PublicFolderURL,LastPost)
	dim modified,title,hreflink,content,author
	if SyncType = "Full" then
		Set xnEntryNodes = xdXmlDocument.selectNodes("rdf:RDF/item")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename
					case "author" author = xnEntryDet.text
					case "dc:creator" author = xnEntryDet.text
					case "guid" id = xnEntryDet.text
					case "dc:date" created =  xnEntryDet.text
					case "pubDate" created =  xnEntryDet.text
					case "description" content = xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    content = xnEntryDet.xml	
						       end if
					case "dc:title" title = xnEntryDet.text
					case "title" title = xnEntryDet.text
					case "link" hreflink = xnEntryDet.text
				end select
			next
			if instr(xnEntrynodes(i).xml,"rdf:about") then
				id = xnEntrynodes(i).attributes.getNamedItem("rdf:about").nodeValue
			end if
			call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
		next
	else 
		Set xnEntryNodes = xdXmlDocument.selectNodes("rdf:RDF/item")
		for i = xnEntryNodes.length-1 to 0 step -1
			set xnEntryDets = xnEntrynodes(i).childnodes
			modified = ""
			title = ""
			hreflink = ""
			content = ""
			author = ""
			id = ""
			for each xnEntryDet in xnEntryDets
				select case xnEntryDet.nodename
					case "author" author = xnEntryDet.text
					case "dc:creator" author = xnEntryDet.text
					case "guid" id = xnEntryDet.text
					case "dc:date" created =  xnEntryDet.text
					case "pubDate" created =  xnEntryDet.text
					case "description" content = xnEntryDet.text
					case "content" if instr(xnEntryDet.xml,"mode=""escaped""") then
						            content = xnEntryDet.text
						       else
							    content = xnEntryDet.xml	
						       end if
					case "dc:title" title = xnEntryDet.text
					case "title" title = xnEntryDet.text
					case "link" hreflink = xnEntryDet.text
				end select
			next
			if instr(xnEntrynodes(i).xml,"rdf:about") then
				id = xnEntrynodes(i).attributes.getNamedItem("rdf:about").nodeValue
			end if
			if created <> "" then
				createdcvn = anumit(left(created,20)) 
			end if
			if createdcvn > lastpost then
				call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
				logfile.writeline "post added " & created & "  " & lastpost
			else
				if created = "" then 
					call CreatePost(PublicFolderURL,modified,title,hreflink,content,author,id)
				else
					logfile.writeline "post not added " & created & "  " & lastpost
				end if
			end if
		next
	
	end if
	logfile.writeline "RDF Feed"
end Sub

function CreatePost(PubFolderURL,modified,title,hreflink,content,author,id)

set msgobj = createobject("CDO.Message")
if id = "" then id = title
emlurl = PubFolderURL & "/" & replace(escape(id),"/","") & ".eml"
logfile.writeline emlurl
on error resume next
msgobj.datasource.open emlurl,conn,3,-1
if err.number = 0 then
	on error goto 0
	header =  "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ""http://www.w3.org/TR/html4/strict.dtd"">"
	footer = "<BR>" & "Feed Source : <A href=""" & hreflink & """>" & hreflink & "</a>" & "<BR><BR>"
	body = header & content & footer
	if body <> msgobj.htmlbody then
		msgobj.from = author
		msgobj.subject = title
		msgobj.htmlbody = body
		msgobj.fields("http://schemas.exdevblog.com/updated") = modified
		msgobj.fields("urn:schemas:httpmail:read") = false
		msgobj.fields.update
		msgobj.datasource.save
		logfile.writeline "Modified"
		set msgobj = nothing
	else
		logfile.writeline "Message Exist no content changes"
	end if
else
	on error goto 0
	msgobj.from = author
	msgobj.subject = title
	if id = "" then id = title
	header =  "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ""http://www.w3.org/TR/html4/strict.dtd"">"
	footer = "<BR>" & "Feed Source : <A href=""" & hreflink & """>" & hreflink & "</a>" & "<BR><BR>"
	msgobj.htmlbody = header & content & footer
	msgobj.fields("http://schemas.exdevblog.com/updated") = modified
	msgobj.fields("urn:schemas:httpmail:read") = false
	msgobj.fields.update
	msgobj.datasource.saveto emlurl,conn,3,67108864
	logfile.writeline "Created"
end if
set msgobj = nothing

end function


function ModifyPost(PubFolderURL,modified,title,hreflink,content,author,id)

if id = "" then id = title
set msgobj = createobject("CDO.Message")
emlurl = PubFolderURL & "/" & replace(escape(id),"/","") & ".eml"
msgobj.datasource.open emlurl,conn,3
header =  "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ""http://www.w3.org/TR/html4/strict.dtd"">"
footer = "<BR>" & "Feed Source : <A href=""" & hreflink & """>" & hreflink & "</a>" & "<BR><BR>"
msgobj.htmlbody = header & content & footer
msgobj.fields("http://schemas.exdevblog.com/updated") = modified
msgobj.fields("urn:schemas:httpmail:read") = false
msgobj.fields.update
logfile.writeline emlurl
msgobj.datasource.save
logfile.writeline "Modified Post"
set msgobj = nothing

end function

user = wscript.arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\" & wscript.arguments(0) & ".csv"
set wfile = fso.opentextfile(fname,2,true)
wfile.writeline("User,Subject,UTC timestart,UTC timeend, InstanceType, OutlookTimezone,CDOTimezoneEnumID")
public datefrom
public dateto
datefrom = "2006-03-26T10:00:00Z"
dateto = "2006-04-02T10:00:00Z"
domainname = "domain.com"
sConnString = "file://./backofficestorage/" & domainname
sConnString = sConnString & "/mbx/" & user & "/calendar"
call QueryCalendarFolder(sConnString,user)

wscript.echo "Done"


Public Sub QueryCalendarFolder(sConnString,user)
SSql = "SELECT ""DAV:href"", ""DAV:parentname"", ""urn:schemas:calendar:timezoneid"", ""urn:schemas:httpmail:subject"", "
SSql = SSql & """http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234"", ""urn:schemas:calendar:timezone"", "
SSql = SSql & """urn:schemas:calendar:instancetype"", ""urn:schemas:calendar:dtstart"" , ""urn:schemas:calendar:dtend"" "
SSql = SSql & "FROM scope('shallow traversal of """ & sConnString & """') " 
SSql = SSql & " Where ""DAV:isfolder"" = false AND ""DAV:ishidden"" = false "           
SSql = SSql & "AND ""urn:schemas:calendar:dtend"" > CAST(""" & datefrom & """ as 'dateTime') " _
& "AND ""urn:schemas:calendar:dtstart"" < CAST(""" & dateto & """ as 'dateTime')"  
   Set oConn = CreateObject("ADODB.Connection")
   oConn.Provider = "Exoledb.DataSource"
   oConn.Open sConnString
   Set oRecSet = CreateObject("ADODB.Recordset")
   oRecSet.CursorLocation = 3
   oRecSet.Open sSQL, oConn.ConnectionString
   if err.number <> 0 then wfile.writeline(user & "," & "Error Connection to Mailbox")
   While oRecSet.EOF <> True
      select case oRecSet.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value
	case "(GMT+10:00) Canberra, Melbourne, Sydney" susapt = 1
						       upto = "(GMT+10:00) Canberra, Melbourne, Sydney (Commonwealth Games)"
	case "(GMT+09:30) Adelaide" susapt = 1
				    upto = "(GMT+09:30) Adelaide (Commonwealth Games)"
	case "(GMT+10:00) Hobart" susapt = 1
				  upto = "(GMT+10:00) Hobart (Commonwealth Games)" 
	case "(GMT+10:00) Canberra, Melbourne, Sydney (Commonwealth Games)" susapt = 2
	case "(GMT+09:30) Adelaide (Commonwealth Games)" susapt = 2
	case "(GMT+10:00) Hobart (Commonwealth Games)" susapt = 2
      end select
      select case oRecSet.fields("urn:schemas:calendar:timezoneid").value
        case "78" susapt = 2
	case "79" susapt = 2
	case "80" susapt = 2
	case "57" susapt = 1
		  uptoid = "78"
	case "19" susapt = 1
		  uptoid = "79"
	case "42" susapt = 1
		  uptoid = "80"
      end select
      if susapt = 0 then  
         if instr(oRecSet.fields("urn:schemas:calendar:timezone").value,"TZID:GMT +0930") then 
		susapt = 1
		upto = "(GMT+09:30) Adelaide (Commonwealth Games)"
	 end if
         if instr(oRecSet.fields("urn:schemas:calendar:timezone").value,"TZID:GMT +1000") then 
		susapt = 1
		upto = "(GMT+10:00) Canberra, Melbourne, Sydney (Commonwealth Games)"
	 end if
	
      end if
      if susapt = 1 then
      		Wscript.echo User
      		wscript.echo oRecSet.fields("DAV:Href").value
     		wscript.echo oRecSet.fields("urn:schemas:httpmail:subject").value
      		wscript.echo oRecSet.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value
     		wscript.echo oRecSet.fields("urn:schemas:calendar:timezoneid").value
     		wscript.echo oRecSet.fields("urn:schemas:calendar:instancetype").value
     		wscript.echo 
      		wfile.writeline(user & "," & oRecSet.fields("urn:schemas:httpmail:subject").value & ","_
     		& oRecSet.fields("urn:schemas:calendar:dtstart").value & "," & oRecSet.fields("urn:schemas:calendar:dtend").value _
     		& "," & oRecSet.fields("urn:schemas:calendar:instancetype").value  & "," _
     		& replace(replace(oRecSet.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value,vbcrlf,""),",","") _
     		& "," & oRecSet.fields("urn:schemas:calendar:timezoneid").value)  
	      set apptobj = createobject("ADODB.Record")
      	      apptobj.open cstr(oRecSet.fields("DAV:HREF").value),,3
      	      wscript.echo dateadd("h",-1,oRecSet.Fields("urn:schemas:calendar:dtstart").value)
     	      apptobj.Fields("urn:schemas:calendar:dtstart").value = dateadd("h",-1,oRecSet.Fields("urn:schemas:calendar:dtstart").value)
      	      apptobj.Fields("urn:schemas:calendar:dtend").value = dateadd("h",-1,oRecSet.Fields("urn:schemas:calendar:dtend").value)
      	      if uptoid <> "" then apptobj.Fields("urn:schemas:calendar:timezoneid").value = uptoid
	      if upto <> "" then apptobj.Fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value = upto
	      apptobj.Fields.update
     	      apptobj.close
	      uptoid = ""
	      upto = ""
      end if
	susapt = 0
	oRecSet.MoveNext
   wend
   oRecSet.Close
   oConn.Close
   Set oRecSet = Nothing
   Set oConn = Nothing
End Sub



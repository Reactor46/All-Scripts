connstr = "Provider=SQLOLEDB.1;Data Source=" & SQLSVR & ";Initial Catalog=" & DB & ";Trusted_Connection=Yes;"
Set myConn = CreateObject("ADODB.Connection")
Set myCommand = CreateObject("ADODB.Command" )
Set rs = CreateObject("ADODB.Recordset")
myConn.Open connstr
Set myCommand.ActiveConnection = myConn
dim lastMod, nextRun, reqTime, fmtDT

rs.Open "SELECT ModifiedBy as modBy FROM Settings WHERE Name = 'RunScheduled'", myConn
lastMod = rs("modBy")
rs.close
rs.Open "SELECT Value as time FROM Settings WHERE Name = 'RunScheduled'", myConn
reqTime = rs("time")
rs.close
nextRun = DateAdd("d",1,Date()) & " 10:00:00 AM"
fmtDT = FormatDateTime(reqTime)

IF DateDiff("s",Now, fmtDT) < 0 THEN

IF Time < "09:55:59 AM" THEN
    nextRun = DateAdd("d",0,Date()) & " 09:55:59 AM"
ELSE 
    nextRun = DateAdd("d",1,Date()) & " 09:55:59 AM"
END IF

 myCommand.CommandText = "UPDATE Settings SET Value = '" & Now & "', ModifiedBy = '" & lastMod & "' WHERE Name = 'ReportRunTime'"
 myCommand.Execute
 myCommand.CommandText = "UPDATE Settings SET Value = '" & nextRun & "', ModifiedBy = 'System' WHERE Name = 'RunScheduled'"
 myCommand.Execute
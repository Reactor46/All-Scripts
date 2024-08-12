'*******************************************************
Dim server, database, uid, pwd
server = "SERVERNAME"
database = "DATABASE"
'connstr = "Provider=SQLOLEDB.1;Data Source=" & server & ";Initial Catalog=" & database & ";user id = '" & uid & "';password='" & pwd & "'"

connstr = "Provider=SQLOLEDB.1;Data Source=" & server & ";Initial Catalog=" & database & ";Trusted_Connection=Yes;"
Set myConn = CreateObject("ADODB.Connection")
Set myCommand = CreateObject("ADODB.Command" )
myConn.Open connstr
Set myCommand.ActiveConnection = myConn
myCommand.CommandText = "UPDATE Settings SET Value = '" & Now & "' WHERE Name = 'Report-DateTime'"
myCommand.Execute
myConn.Close
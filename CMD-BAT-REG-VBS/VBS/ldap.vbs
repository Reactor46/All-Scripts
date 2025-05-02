Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

objCommand.CommandText = "<GC://<LDAPServer>:389/DC=<domainComponent>,DC=<domainComponent>;(&(objectClass=posixGroup)(memberUid=test1));" & "cn;subtree" 
 
Set objRecordSet = objCommand.Execute

While Not objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("cn")
    Wscript.Echo VbCrLf
    objRecordSet.MoveNext
Wend

objConnection.Close

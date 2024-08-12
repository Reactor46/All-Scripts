'This script will list all disable accounts in the domain
'Cheyenne Harden 9.28.06

Const ADS_UF_ACCOUNTDISABLE = 2
Const OPEN_FILE_FOR_WRITING = 2
strFile = "disabled.txt"
strWritePath = "c:\" & strFile
strDirectory = "c:\"

'#########
Set objFSO1 = CreateObject("Scripting.FileSystemObject")

If objFSO1.FileExists("c:\" & strFile) Then
	Set objFolder = objFSO1.GetFile("c:\" & strFile)

Else
	Set objFile = objFSO1.CreateTextFile(strDirectory & strFile)
	'Wscript.Echo "Just created " & objFolder & "\" & strFile
	objFile = ""

End If
'#########
Set fso = CreateObject("Scripting.FileSystemObject")
Set textFile = fso.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
' Below put in your domain name like this (e.g., dc=CONTOSO,dc=com)
' The 2nd dc= could be "com, net, org, local" or what ever you use.
objCommand.CommandText = _
"<GC://dc=YOURDOMAINHERE,dc=YOURDOMAINSUFFIXHERE>;(objectCategory=User)" & _
";userAccountControl,distinguishedName;subtree" 
Set objRecordSet = objCommand.Execute

intCounter = 0
While Not objRecordset.EOF
intUAC=objRecordset.Fields("userAccountControl")
If intUAC AND ADS_UF_ACCOUNTDISABLE Then
'WScript.echo objRecordset.Fields("distinguishedName") & " is disabled"
textFile.WriteLine(objRecordset.Fields("distinguishedName"))
intCounter = intCounter + 1
End If
objRecordset.MoveNext
Wend

WScript.Echo VbCrLf & "A total of " & intCounter & " accounts are disabled."


objConnection.Close

WScript.Echo "Done..."
AppToRun = "c:\disabled.txt"
CreateObject("Wscript.Shell").Run AppToRun
WScript.Quit
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strTemp = objshell.ExpandEnvironmentStrings("%TEMP%") & "\"		' Get path to user profile's "temp" folder
htaFilePath = strTemp & "staffdetails.hta"

Set htmlFile = objFSO.CreateTextFile(htaFilePath, TRUE)	' Create new text file in user's temp folder with HTA extension

htmlfile.WriteLine("<HTML>")							' ...and write the HTA content, line-by-line.
htmlfile.WriteLine("<TITLE>Staff Contact Details</TITLE>")
htmlfile.WriteLine("<HEAD>")
htmlfile.WriteLine("<HTA:APPLICATION")
htmlfile.WriteLine("MAXIMIZEBUTTON=""no""")
htmlfile.WriteLine("SINGLEINSTANCE=""yes"">")
htmlfile.WriteLine("SYSMENU=""no"">")
htmlfile.WriteLine("<STYLE type=""text/css"">")			' Change or remove this <STYLE> block if you so wish
htmlfile.WriteLine("")									' I just like to keep things easy to read on-screen
htmlfile.WriteLine("#TABLE1")
htmlfile.WriteLine("{")
htmlfile.WriteLine("	font-family: ""Lucida Sans Unicode"", ""Lucida Grande"", Sans-Serif;")
htmlfile.WriteLine("	font-size: 14px;")
htmlfile.WriteLine("	background: #fff;")
htmlfile.WriteLine("	margin: 25px;")
htmlfile.WriteLine("	width: 95%;")
htmlfile.WriteLine("	border-collapse: collapse;")
htmlfile.WriteLine("	text-align: left;")
htmlfile.WriteLine("}")
htmlfile.WriteLine("#TABLE1 TH")
htmlfile.WriteLine("{")
htmlfile.WriteLine("	font-size: 16px;")
htmlfile.WriteLine("	font-weight: normal;")
htmlfile.WriteLine("	color: #039;")
htmlfile.WriteLine("	padding: 10px 8px;")
htmlfile.WriteLine("	border-bottom: 2px solid #6678b1;")
htmlfile.WriteLine("}")
htmlfile.WriteLine("#TABLE1 TD")
htmlfile.WriteLine("{")
htmlfile.WriteLine("	border-bottom: 1px solid #ccc;")
htmlfile.WriteLine("	color: #669;")
htmlfile.WriteLine("	padding: 6px 8px;")
htmlfile.WriteLine("}")
htmlfile.WriteLine("</STYLE>")
htmlfile.WriteLine("<SCRIPT Language=""VBScript"">")
htmlfile.WriteLine("	 Sub Window_Onload")
htmlfile.WriteLine("		window.resizeTo 700,700")	' Resize window to preferred dimensions
htmlfile.WriteLine("	 End Sub")
htmlfile.WriteLine("</SCRIPT>")
htmlfile.WriteLine("</HEAD>")
htmlfile.WriteLine("")
htmlfile.WriteLine("<BODY><CENTER>")
htmlfile.WriteLine("<H2 style=""font-family: Lucida Sans Unicode; color: #6678b1"">Staff Contact Details</H2>")
htmlfile.WriteLine("<TABLE id=""table1"" border=""0""><TBODY>")
htmlfile.WriteLine("<THEAD>")
htmlfile.WriteLine("	<TR>")
htmlfile.WriteLine("		<TH>Name</TH>")					' Column headings
htmlfile.WriteLine("		<TH>Email Address</TH>")
htmlfile.WriteLine("		<TH>Mobile Phone Number</TH>")
htmlfile.WriteLine("		</TR>")
htmlfile.WriteLine("</THEAD>")
htmlfile.WriteLine("")
htmlfile.WriteLine("<TBODY>")

Set objContainer = GetObject("LDAP://ou=Branch 1,dc=aldergrovecu,dc=local")	' Change to match the OU you want to list
objContainer.Filter = Array("user")
For Each objChild In objContainer
		If objChild.mail <> "" Then		' Report only the users with email addresses on their AD profile; change this as per your requirement
		htmlfile.WriteLine("<TR>")
		htmlfile.WriteLine("	<TD>" & objChild.FullName & "</TD>")	' On each table row, write the ADUC details
		htmlfile.WriteLine("	<TD>" & objChild.mail & "</TD>")		' of each user to match the column headings
		htmlfile.WriteLine("	<TD>" & objChild.Mobile & "</TD>")
		htmlfile.WriteLine("</TR>")
		End If
Next

htmlfile.WriteLine("</TBODY>")
htmlfile.WriteLine("</TABLE>")
htmlfile.WriteLine("</CENTER>")
htmlfile.WriteLine("</BODY></HTML>")
objShell.Run htaFilePath	'Open the HTA file to display the results


set objFSO = Nothing		' Housekeeping
set objShell = Nothing
set objContainer = Nothing
On Error Resume Next
Const ForAppending = 8
Const HARD_DISK = 3
Const ForReading = 1

'Declaring the variables
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set SrvList = objFSO.OpenTextFile("Server_List.txt", ForReading)
Set ReportFile = objFSO.OpenTextFile ("Diskspace_status.html", ForAppending, True)
i = 0

'Initializing the HTML Tags for better formatting
ReportFile.writeline("<html>")
ReportFile.writeline("<head>")
ReportFile.writeline("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>")
ReportFile.writeline("<title>" & "SQL Servers Disk Space Report</title>")
ReportFile.writeline("<style type='text/css'>")
ReportFile.writeline("<!--")
ReportFile.writeline("td {")
ReportFile.writeline("font-family: Tahoma;")
ReportFile.writeline("font-size: 11px;")
ReportFile.writeline("border-top: 1px solid #999999;")
ReportFile.writeline("border-right: 1px solid #999999;")
ReportFile.writeline("border-bottom: 1px solid #999999;")
ReportFile.writeline("border-left: 1px solid #999999;")
ReportFile.writeline("padding-top: 0px;")
ReportFile.writeline("padding-right: 0px;")
ReportFile.writeline("padding-bottom: 0px;")
ReportFile.writeline("padding-left: 0px;")
ReportFile.writeline("}")
ReportFile.writeline("body {")
ReportFile.writeline("margin-left: 5px;")
ReportFile.writeline("margin-top: 5px;")
ReportFile.writeline("margin-right: 0px;")
ReportFile.writeline("margin-bottom: 10px;")
ReportFile.writeline("")
ReportFile.writeline("table {")
ReportFile.writeline("border: thin solid #000000;")
ReportFile.writeline("}")
ReportFile.writeline("-->")
ReportFile.writeline("</style>")
ReportFile.writeline("</head>")
ReportFile.writeline("<body>")

ReportFile.writeline("<table width='50%'>")
ReportFile.writeline("<tr bgcolor='#CCCCCC'>")
ReportFile.writeline("<td colspan='7' height='25' align='center'>")
ReportFile.writeline("<font face='tahoma' color='#003399' size='2'><strong>SQL Servers Disk Space Report</strong></font>")
ReportFile.writeline("</td>")
ReportFile.writeline("</tr>")
ReportFile.writeline("</table>")


'Declaring the Server Name for report generation
Do Until SrvList.AtEndOfStream
	StrComputer = SrvList.Readline

	ReportFile.writeline("<table width='50%'><tbody>")
	ReportFile.writeline("<tr bgcolor='#CCCCCC'>")
	ReportFile.writeline("<td width='50%' align='center' colSpan=6><font face='tahoma' color='#003399' size='2'><strong>" & StrComputer & "</strong></font></td>")
	ReportFile.writeline("</tr>")


	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")

	ReportFile.writeline("<tr bgcolor=#CCCCCC>")
		ReportFile.writeline("<td width='05%' align='center'>Drive / Mount</td>")
		ReportFile.writeline("<td width='05%' align='center'>Total Capacity (in GB)</td>")
		ReportFile.writeline("<td width='05%' align='center'>Used Capacity (in GB)</td>")
		ReportFile.writeline("<td width='05%' align='center'>Free Space (in GB)</td>")
		ReportFile.writeline("<td width='05%' align='center'>Freespace %</td>")
	ReportFile.writeline("</tr>")

		'Starting the loop to gather values from all Hard Drives
		For Each objDisk in colDisks

			'Delcaring the Variables

			TotSpace=Round(((objDisk.Size)/1073741824),2)
			FrSpace=Round(objDisk.FreeSpace/1073741824,2)
			FrPercent=Round((FrSpace / TotSpace)*100,0)
			UsSpace=Round((TotSpace - FrSpace),2)
			Drv=objDisk.DeviceID
			VolName=objDisk.DeviceID

			'Lnt=Len(VolName)

			'If  Len(VolName) =  3 then
				If FrPercent > 20 Then
					ReportFile.WriteLine "<tr><td align=center>" & Drv & "</td><td align=center>" & TotSpace & "</td><td align=center>" & UsSpace & "</td><td align=center>" & FrSpace & "</td><td BGCOLOR='#00FF00' align=center>" & FrPercent & "%" &"</td></tr>"
				ElseIf FrPercent < 10 Then
					ReportFile.WriteLine "<tr><td align=center>" & Drv & "</td><td align=center>" & TotSpace & "</td><td align=center>" & UsSpace & "</td><td align=center>" & FrSpace & "</td><td bgcolor='#FF0000' align=center>" & FrPercent & "%" &"</td></tr>"
				Else
					ReportFile.WriteLine "<tr><td align=center>" & Drv & "</td><td align=center>" & TotSpace & "</td><td align=center>" & UsSpace & "</td><td align=center>" & FrSpace & "</td><td bgcolor='#FBB917' align=center>" & FrPercent & "%" &"</td></tr>"
				End If
			'Else
			'End If
		Next

	ReportFile.writeline("<tr>")
	ReportFile.writeline("<td width='50%' colSpan=6>&nbsp;</td>")
	ReportFile.writeline("</tr>")

	ReportFile.writeline("</tbody></table>")
Loop
ReportFile.WriteLine "</body></html>"
On Error Resume Next
Dim sqlRS
Dim ADCon, ocommand, strsql 	'As ADODB.Command
Dim RetCollectionID, RetColTableName, RetOfferID, RetPolicyID, RetPolicyVersion
Dim StrSendText, ErrCount, SysResFound, StrSysResListArr()
Dim RetHTTPStatus
Const ForAppend = 8, adOpenStatic = 3, adLockOptimistic = 3
Const ForReading = 1, ForWriting = 2, ForAppending = 8

StrDateArry = Split(FormatDateTime(NOW,vbShortDate),"/")
StrYear = StrDateArry(2)
If Len(StrDateArry(0)) < 2 Then
	StrMonth = "0" & StrDateArry(0)
Else
	StrMonth = StrDateArry(0)
End If
If Len(StrDateArry(1)) < 2 Then
	StrDay = "0" & StrDateArry(1)
Else
	StrDay = StrDateArry(1)
End If

StrTimeArry = Split(FormatDateTime(NOW,vbShortTime),":")
StrTime = StrTimeArry(0) & StrTimeArry(1)
StrTodayDateTime = StrYear & StrMonth & StrDay & StrTime

Set oFileSys = CreateObject("Scripting.FileSystemObject")
Set oNetwork = CreateObject("Wscript.Network")
StrScriptPath = Left(wscript.ScriptFullName,InStrRev(wscript.ScriptFullName,"\"))
StrOutputResult = StrScriptPath & "ActivityLog\" & StrTodayDateTime & "-SLP-MP-Activity.log"
StrSummaryHTM = StrScriptPath & "SummaryHTM\" & StrTodayDateTime & "-SLP-MP-STATUS.HTM"

TotalSysListCount = 0
ErrCount = 0
Total_ErrCount = 0
'##################################################################################
'##################################################################################
'************Replace the DATABASE-SERVER-NAME with your actual DB server name
'************Replace the SMS_SiteCode with your site code SMS DataBase Name
'##################################################################################
'##################################################################################
StrSMSDBServer = "DATABASE-SERVER-NAME"
StrSiteDB = "SMS_SiteCode"
sComputerName = oNetwork.ComputerName

Set StrWriteOutputResults = oFileSys.CreateTextFile(StrOutputResult,ForWriting)
Set StrOutputHTML = oFileSys.CreateTextFile(StrSummaryHTM,ForWriting)

StrWriteOutputResults.WriteLine "Start retrieving SMS SLP & MP Information < " & Now & " >"

StrWriteOutputResults.WriteLine		
StrWriteOutputResults.WriteLine "Connecting to SMS Database ...................................................."
StrWriteOutputResults.WriteLine "Database Server : " & StrSMSDBServer
StrWriteOutputResults.WriteLine "Database Name : " & strSiteDB
StrWriteOutputResults.WriteLine

Set oSMSConn = CreateObject("ADODB.Connection")
Set sqlRS = CreateObject("ADODB.Recordset")

On Error Resume Next
oSMSConn.Open "Provider = SQLOLEDB;" & "Data Source =" & StrSMSDBServer & ";"&_
	    "Initial Catalog=" & strSiteDB & ";"&_
	    "INTEGRATED SECURITY=sspi;"
ErrNumber = Err.Number
ErrDescription = Err.Description
Err.Clear
On Error Goto 0
If ErrNumber <> 0 Then
	StrWriteOutputResults.WriteLine "Error encountered while connecting to Database, Error: " & ErrNumber & ", Description: " & ErrDescription
	StrWriteOutputResults.WriteLine 
	StrWriteOutputResults.WriteLine  "Query Ended < " & Now & " >"
	StrWriteOutputResults.Close
	Wscript.Quit
End If

GetSMSSLP_MPInfo ' Start Get SLP & MP Information from SMS SQL DB

'### Start Creating the HTM Summary Header & Table ###
StrHTMLBodyText = "<HTML>" & VBCR
StrHTMLBodyText = StrHTMLBodyText  & "<BODY>" & VBCR
StrHTMLBodyText = StrHTMLBodyText & "<TABLE id=" & Chr(34) & "Table1" & Chr(34) & " cellSpacing=" & Chr(34) & "1" & Chr(34) & _
		" cellPadding=" & Chr(34) & "1" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & "> " & VBCR
StrHTMLBodyText = StrHTMLBodyText & "<TR><TH bgcolor=" & Chr(34) & "#33CCFF" & Chr(34) & "><a href=" & Chr(34) & "file://\\" & _
		sComputerName & "\ChkMPSLPLog$\" & StrTodayDateTime & "-SLP-MP-Activity.log" & Chr(34) & ">SMS Management Point and SLP Healthcheck</a></TH></TR> " & VBCR
StrHTMLBodyText = StrHTMLBodyText & "</FONT> </TD></TR></TABLE><br> " & VBCR
StrHTMLBodyText = StrHTMLBodyText & "<TABLE id=" & Chr(34) & "TitleTable" & Chr(34) & " cellSpacing=" & Chr(34) & "1" & Chr(34) & _
		" cellPadding=" & Chr(34) & "1" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & "> " & VBCR
StrHTMLBodyText = StrHTMLBodyText & "<TR><TH> Server Name</TH><TH> Site Code </TH><TH> Server Role </TH><TH> Status </TH>" & _
		"<TH> Checked Time </TH></TR> "

QuerySLP_MPStatus StrSysResListArr ' Read Start Get SLP & MP status

StrHTMLBodyText = StrHTMLBodyText & "</TABLE><PRE></PRE></BODY></HTML>"

StrOutputHTML.WriteLine StrHTMLBodyText

SendMailAlerts StrHTMLBodyText,TotalSysListCount

StrWriteOutputResults.WriteLine  "Query Ended < " & Now & " >"
StrWriteOutputResults.WriteLine 
StrWriteOutputResults.WriteLine "< Summary > Total SMS System Roles = " & TotalSysListCount & " , " & "Total Error = " & Total_ErrCount
StrOutputHTML.Close
StrWriteOutputResults.Close

'Wscript.Echo "Query completed !!!!"
Wscript.Quit

'######### Get SMS Server Located Point & Management Points from SQL DB ####################
Sub GetSMSSLP_MPInfo
Dim StrSiteCode, StrServerName, StrRoleName
SysResFound = 0
strSQL = "Select SiteCode, ServerName, RoleName from sysreslist where ( rolename = 'SMS Management Point' or " & _
	"Rolename = 'SMS Server Locator Point') Order by SiteCode"

'open cx to SMS SQL
On Error Resume Next
sqlRS.Open strSql, oSMSConn, adOpenStatic, adLockOptimistic
ErrNumber = Err.Number
ErrDescription = Err.Description
Err.Clear
On Error Goto 0
If ErrNumber <> 0 Then
	StrWriteOutputResults.WriteLine "Error encountered while query database for SysResList information, Error: " & ErrNumber & ", Description: " & ErrDescription
	StrWriteOutputResults.WriteLine 
	StrWriteOutputResults.WriteLine  "Query Ended < " & Now & " >"
	Wscript.Quit
End If

StrWriteOutputResults.WriteLine "Beginning SysResList information query from SMS DB ............................."
StrWriteOutputResults.WriteLine

Do Until sqlRS.EOF
	StrSiteCode = sqlRS("SiteCode")
	StrServerName = sqlRS("ServerName")
	StrRoleName = sqlRS("RoleName")
	StrWriteOutputResults.WriteLine StrSiteCode & "," & StrServerName & "," & StrRoleName
	Err.Clear
   	ReDim Preserve StrSysResListArr(SysResFound)
   	StrSysResListArr(UBound(StrSysResListArr)) = Array(StrSiteCode,StrServerName,StrRoleName)
   	SysResFound = SysResFound + 1
	TotalSysListCount = SysResFound
	sqlRS.Movenext
Loop

If SysResFound = 0 Then
	StrWriteOutputResults.WriteLine
	StrWriteOutputResults.WriteLine "No SysResList Found."
	StrWriteOutputResults.WriteLine
End If

StrWriteOutputResults.WriteLine
StrWriteOutputResults.WriteLine "Finish SysResList information query from SMS DB ............................."
StrWriteOutputResults.WriteLine
sqlRS.Close

End Sub


'######################### Get SLP and MP Status #####################################
Sub QuerySLP_MPStatus(SysResListArr)
Dim TempSiteCode
Const ArrSiteCode = 0, ArrServerName = 1, ArrRoleName = 2

StrChkRoleName1 = "SMS Server Locator Point"
StrChkRoleName2 = "SMS Management Point"

StrWriteOutputResults.WriteLine "Beginning Querying SLP & MP Status ....................."
StrWriteOutputResults.WriteLine

'#####################################################################################
'#####################################################################################
'Change SLP IP Adress and IP range details as per your environment - site&ad=<AD Site Name>
'More details in the blog post http://blogs.technet.com/b/manageabilityguys/archive/2011/06/30/testing-slp-availability.aspx
'#####################################################################################
'#####################################################################################

For x = 0 to Ubound(StrSysResListArr)
	SMSStrSiteCode = StrSysResListArr(x)(ArrSiteCode)
	SMSStrServerName = StrSysResListArr(x)(ArrServerName)
	SMSStrRoleName = StrSysResListArr(x)(ArrRoleName)
	If Ucase(SMSStrRoleName) = Ucase(StrChkRoleName1) Then
		RetHTTPStatus = ""
		SlpURL = "http://" & SMSStrServerName & "/sms_slp/slp.dll?site&ad=<AD Site Name>"
		StrWriteOutputResults.WriteLine "Querying " & SMSStrServerName & " for " & StrChkRoleName1
		StrWriteOutputResults.WriteLine "Get HTTP >" & SlpURL
		GetSLPStatus SlpURL,RetHTTPStatus
		CreateHTMTable SMSStrServerName, SMSStrSiteCode, SMSStrRoleName, RetHTTPStatus
	ElseIf Ucase(SMSStrRoleName) = Ucase(StrChkRoleName2) Then
		RetHTTPStatus = ""
		StrWriteOutputResults.WriteLine "Querying " & SMSStrServerName & " for " & StrChkRoleName2
		GetMPStatus SMSStrServerName,RetHTTPStatus
		CreateHTMTable SMSStrServerName, SMSStrSiteCode, SMSStrRoleName, RetHTTPStatus
	End If
Next

StrWriteOutputResults.WriteLine "Finish Querying SLP & MP Status ........................." & Now
StrWriteOutputResults.WriteLine

End Sub

Sub GetSLPStatus(StrSLPURL,StrStatus)
Set objHTTP = CreateObject("msxml2.xmlhttp")
On Error Resume Next
objHTTP.Open "GET",StrSLPURL,False
objHTTP.Send
On Error Goto 0

RetResponseText = Ucase(Trim(objHTTP.ResponseText))
RetStatus = objHTTP.Status

If RetStatus = 200 AND InStr(RetResponseText,"NO RESULTS") = 0 Then
	RetHTTPStatus = "GREEN"
	StrWriteOutputResults.WriteLine "Status > " & RetHTTPStatus
	StrWriteOutputResults.WriteLine "ResponseText > " & RetResponseText
	StrWriteOutputResults.WriteLine
Else
	RetHTTPStatus = "RED"
	StrWriteOutputResults.WriteLine "Status > " & RetHTTPStatus
	StrWriteOutputResults.WriteLine "ResponseText > " & RetResponseText
	StrWriteOutputResults.WriteLine
	Total_ErrCount = Total_ErrCount + 1
End if

End Sub

'############# Get MPList, MPCert & MPPolicy ##########################
'######################################################################
'Policy ID ("C0120820-C0100265-2D40483D.1_00") mentioned in the below line is dummy one. 
'Use PolicySpy tool to findout any of the existing policyID in your environment and use the same for policy test.
'######################################################################
Sub GetMPStatus(StrMPServer,StrStatus)
Const SoftwarePolicyID = "C0120820-C0100265-2D40483D.1_00" 'Sample software policy created for healthcheck
Dim GreenCount

MPListURL =   "http://" & StrMPServer & "/SMS_MP/.sms_aut?mplist"
MPCertURL =   "http://" & StrMPServer & "/SMS_MP/.sms_aut?mpcert"
MPPolicyURL = "http://" & StrMPServer & "/SMS_MP/.sms_pol?" & SoftwarePolicyID
GreenCount = 0

Set objHTTP = CreateObject("msxml2.xmlhttp")
On Error Resume Next
StrWriteOutputResults.WriteLine "Get HTTP >" & MPListURL
objHTTP.Open "GET",MPListURL,False
objHTTP.Send
On Error Goto 0

If objHTTP.Status = 200 Then
	StrWriteOutputResults.WriteLine "Status > " & "GREEN"
	StrWriteOutputResults.WriteLine
	GreenCount = GreenCount + 1
Else
	StrWriteOutputResults.WriteLine "Status > " & "RED"
	StrWriteOutputResults.WriteLine
	Total_ErrCount = Total_ErrCount + 1
End if

On Error Resume Next
StrWriteOutputResults.WriteLine "Get HTTP >" & MPCertURL
objHTTP.Open "GET",MPCertURL,False
objHTTP.Send
On Error Goto 0

If objHTTP.Status = 200 Then
	StrWriteOutputResults.WriteLine "Status > " & "GREEN"
	StrWriteOutputResults.WriteLine
	GreenCount = GreenCount + 1
Else
	StrWriteOutputResults.WriteLine "Status > " & "RED"
	StrWriteOutputResults.WriteLine
	Total_ErrCount = Total_ErrCount + 1
End if

On Error Resume Next
StrWriteOutputResults.WriteLine "Get HTTP >" & MPPolicyURL
objHTTP.Open "GET",MPPolicyURL,False
objHTTP.Send
On Error Goto 0

If objHTTP.Status = 200 Then
	StrWriteOutputResults.WriteLine "Status > " & "GREEN"
	StrWriteOutputResults.WriteLine
	GreenCount = GreenCount + 1
Else
	StrWriteOutputResults.WriteLine "Status > " & "RED"
	StrWriteOutputResults.WriteLine
	Total_ErrCount = Total_ErrCount + 1
End if

If GreenCount = 3 Then
	RetHTTPStatus = "GREEN"
ElseIf GreenCount = 2 or GreenCount = 1 Then
	RetHTTPStatus = "YELLOW"
ElseIf GreenCount = 0 Then
	RetHTTPStatus = "RED"
End If

End Sub


'######## Create Summary HTM Report for SLP and MP Status  ###########
Sub CreateHTMTable(StrServerName, StrSiteCode, StrRoleName, StrStatusCode)

If StrStatusCode = "GREEN" Then
	StrHTMLBodyText = StrHTMLBodyText & "<TR><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrServerName & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & _
		StrSiteCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrRoleName & "</TD>" & "<TD bgcolor=" & Chr(34) & "#33CC00" & Chr(34) & _
		"><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrStatusCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & Now & "</TD></TR>"
ElseIf StrStatusCode = "YELLOW" Then
	StrHTMLBodyText = StrHTMLBodyText & "<TR><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrServerName & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & _
		StrSiteCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrRoleName & "</TD>" & "<TD bgcolor=" & Chr(34) & "#FFFF33" & Chr(34) & _
		"><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrStatusCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & Now & "</TD></TR>"
ElseIf StrStatusCode = "RED" Then
	StrHTMLBodyText = StrHTMLBodyText & "<TR><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrServerName & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & _
		StrSiteCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrRoleName & "</TD>" & "<TD bgcolor=" & Chr(34) & "#FF0033" & Chr(34) & _
		"><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & StrStatusCode & "</TD><TD><CENTER><FONT size=" & Chr(34) & "2" & Chr(34) & ">" & Now & "</TD></TR>"

End If

End Sub


'##################### Send Email Alerts ##############################################################
'######################################################################################################
'######################################################################################################
'Change your Email Setting as per the Organiztion 
'Configure SMTP Server Name, FROM and TO MailIDs
'######################################################################################################
'######################################################################################################
'######################################################################################################
Function SendMailAlerts(StrSummaryHTML,TotalSysCount)
RecipientList ="team@CompanyName.com"
smtpsrv = "smtp.CompanyName.intra/com"

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "noreply@CompanyName.com"
objEmail.To = RecipientList
objEmail.Subject = "SLP_MP SCCM Infrastructure Management " & Now & " Total = " & TotalSysCount & " Error = " & Total_ErrCount
objEmail.HTMLBody = StrSummaryHTML
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpsrv 
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send

End Function

'##################### End Send Email Alerts ##############################################################


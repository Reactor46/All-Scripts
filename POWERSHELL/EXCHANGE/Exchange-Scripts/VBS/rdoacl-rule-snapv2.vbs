snServername = wscript.arguments(0)
mbMailboxName = wscript.arguments(1)

unDisplayname = ""

csCurrentSnapFileName = "c:\temp\currentSnap-" & mbMailboxName & ".xml"
psPreviousSnapFileName = "c:\temp\prevSnap-" & mbMailboxName & ".xml"
adArchieveDirectory = "c:\temp\SnapArchive\"
rfReportFileName = "c:\temp\ACL-Rule-ChangeReport-" & mbMailboxName 
rrRightReport = 0
Set fso = CreateObject("Scripting.FileSystemObject")
'check SnapArchive'
If Not fso.FolderExists(adArchieveDirectory) Then
	wscript.echo "Archive Folder Created"
	fso.createfolder(adArchieveDirectory)
End if
''

If fso.FileExists(csCurrentSnapFileName) Then
	wscript.echo "Snap Exists"
	If fso.FileExists(psPreviousSnapFileName) Then fso.deletefile(psPreviousSnapFileName)
	fso.movefile csCurrentSnapFileName, psPreviousSnapFileName
	set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
	xdXmlDocument.async="false"
	xdXmlDocument.load(psPreviousSnapFileName)
	Set xnSnaptime = xdXmlDocument.selectNodes("//SnappedACLS")
	For Each exSnap In xnSnaptime
		oldSnap = exSnap.attributes.getNamedItem("SnapDate").nodeValue
		wscript.echo "Snap Taken : " & oldSnap 
		takesnap
		afFileName = adArchieveDirectory & Replace(Replace(Replace(exSnap.attributes.getNamedItem("SnapDate").nodeValue,":",""),",","")," ","") & ".xml"
		wscript.echo "Archiving Old Snap to : " & afFileName
		fso.copyfile psPreviousSnapFileName, afFileName
	Next
	set xdXmlDocument1 = CreateObject("Microsoft.XMLDOM")
	xdXmlDocument1.async="false"
	xdXmlDocument1.load(csCurrentSnapFileName)
	Set ckCurrentPerms = CreateObject("Scripting.Dictionary")
	Set pkPreviousPerms = CreateObject("Scripting.Dictionary")
	Set xnCurrentPermsUsers = xdXmlDocument1.selectNodes("//Folder")
	For Each xnUserNode In xnCurrentPermsUsers
		fnFolderName =  xnUserNode.attributes.getNamedItem("Name").nodeValue
		For Each caACLs In xnUserNode.ChildNodes
			ReDim aclArray1(1)
			aclArray1(0) = caACLs.attributes.getNamedItem("Right").nodeValue
			aclArray1(1) = caACLs.attributes.getNamedItem("Name").nodeValue
			ckCurrentACL = fnFolderName & "|-|" & caACLs.attributes.getNamedItem("User").nodeValue
		    ckCurrentPerms.add   ckCurrentACL, aclArray1
		Next
	Next
	Set xnPrevPermsUsers = xdXmlDocument.selectNodes("//Folder")
	For Each xnUserNode1 In xnPrevPermsUsers
		fnFolderName1 =  xnUserNode1.attributes.getNamedItem("Name").nodeValue
		For Each caACLs1 In xnUserNode1.ChildNodes
			ReDim aclArray1(1)
			aclArray1(0) = caACLs1.attributes.getNamedItem("Right").nodeValue
			aclArray1(1) = caACLs1.attributes.getNamedItem("Name").nodeValue
			pkPrevACL = fnFolderName1 & "|-|" & caACLs1.attributes.getNamedItem("User").nodeValue
		    pkPreviousPerms.add   pkPrevACL, aclArray1
			rem Do a Check for Any Deleted or Changed Permisssions
			If ckCurrentPerms.exists(pkPrevACL) Then
				ckCurrentACLArray = ckCurrentPerms(pkPrevACL)
				If 	ckCurrentACLArray(0) <> caACLs1.attributes.getNamedItem("Right").nodeValue Then
					rrRightReport = 1 
					wscript.echo "Found Changed ACL " 
					wscript.echo "Old Rights : "  & pkPrevACL & "	" & caACLs1.attributes.getNamedItem("Right").nodeValue
					wscript.echo "New Rights : "  & pkPrevACL & "	" & ckCurrentACLArray(0)
					hrmodHtmlReport = hrmodHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & fnFolderName1 & " </font></td>" & vbcrlf
					hrmodHtmlReport = hrmodHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">Old Rights: " &  caACLs1.attributes.getNamedItem("Name").nodeValue _
					& "	" & caACLs1.attributes.getNamedItem("Right").nodeValue &  " </font></td>" & vbcrlf
					hrmodHtmlReport = hrmodHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">New Rights: " _ 
					&  caACLs1.attributes.getNamedItem("Name").nodeValue & "	" & ckCurrentACLArray(0) & " </font></td></tr>" & vbcrlf
				End if
			Else
				rrRightReport = 1 
				hrDelHtmlReport = hrDelHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & fnFolderName1 & " </font></td>" & vbcrlf
				hrDelHtmlReport = hrDelHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  caACLs1.attributes.getNamedItem("Name").nodeValue _
					& "	" & caACLs1.attributes.getNamedItem("Right").nodeValue & " </font></td></tr>" & vbcrlf
				Wscript.echo "Found Deleted ACL : " &  pkPrevACL & "	" & caACLs1.attributes.getNamedItem("Right").nodeValue
			End if
		Next
	Next
	rem Do forward check of ACL's
	For Each dkCurrenPermKey In ckCurrentPerms.keys
		If Not pkPreviousPerms.exists(dkCurrenPermKey) Then
			rrRightReport = 1 
			dkpermsvaluearray = ckCurrentPerms(dkCurrenPermKey)
		    dknewpermarry = Split(dkCurrenPermKey,"|-|")
			hrnewHtmlReport = hrnewHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & dknewpermarry(0) & " </font></td>" & vbcrlf
			hrnewHtmlReport = hrnewHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &   dkpermsvaluearray(1) _
			& "	" & dkpermsvaluearray(0) & " </font></td></tr>" & vbcrlf
			Wscript.echo "Found new ACL : "  & dkCurrenPermKey & "	" &  dkpermsvaluearray(0)
		End if
	Next
	rem Check Rules
	Set ckCurrentRules = CreateObject("Scripting.Dictionary")
	Set pkPreviousRules = CreateObject("Scripting.Dictionary")
	Set xnCurrentRules = xdXmlDocument1.selectNodes("//Rule")
	For Each xnRule In xnCurrentRules
		ReDim ruleArray(1)
		rnRuleName =  xnRule.attributes.getNamedItem("Name").nodeValue
		ruleArray(0) =  xnRule.attributes.getNamedItem("ActionType").nodeValue 
		ruleArray(1) =  xnRule.attributes.getNamedItem("Arg").nodeValue
		ckCurrentRules.add rnRuleName,ruleArray
	Next
	Set xnPrevRules = xdXmlDocument.selectNodes("//Rule")
	For Each xnRule1 In xnPrevRules
		ReDim ruleArray1(1)
		rnRuleName1 =  xnRule1.attributes.getNamedItem("Name").nodeValue
		ruleArray1(0) =  xnRule1.attributes.getNamedItem("ActionType").nodeValue 
		ruleArray1(1) =  xnRule1.attributes.getNamedItem("Arg").nodeValue
		pkPreviousRules.add rnRuleName1, ruleArray1
		rem Do a Check for Any Deleted or Changed rules
		If ckCurrentRules.exists(rnRuleName1) Then
			ckCurrentRuleArray = ckCurrentRules(rnRuleName1)
			If ckCurrentRuleArray(0) <> xnRule1.attributes.getNamedItem("ActionType").nodeValue Then
				rrRightReport = 1 
				wscript.echo "Rule - Action Change"
				wscript.echo "Old Value : " & xnRule1.attributes.getNamedItem("ActionType").nodeValue
				wscript.echo "New Value : " & ckCurrentRuleArray(0)
				hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & rnRuleName1 & " </font></td>" & vbcrlf
			Else
				If ckCurrentRuleArray(1) <> xnRule1.attributes.getNamedItem("Arg").nodeValue Then
					rrRightReport = 1 
					wscript.echo "Rule - Arg Change"
					wscript.echo "Old Value : " & xnRule1.attributes.getNamedItem("Arg").nodeValue
					wscript.echo "New Value : " & ckCurrentRuleArray(1)
					hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & rnRuleName1 & " </font></td>" & vbcrlf
					if ckCurrentRuleArray(0) = 1 Or ckCurrentRuleArray(0) = 2 then 
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  DisplayActionType(xnRule1.attributes.getNamedItem("ActionType").nodeValue) & " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">Old Folder: " &  xnRule1.attributes.getNamedItem("Arg").nodeValue _
						& " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">New Folder: " _ 
						& ckCurrentRuleArray(1) & " </font></td></tr>" & vbcrlf				
					Elseif ckCurrentRuleArray(0) = 6 Or ckCurrentRuleArray(0) = 7 then 
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  DisplayActionType(xnRule1.attributes.getNamedItem("ActionType").nodeValue) & " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">Old Recipients: "  &  xnRule1.attributes.getNamedItem("Arg").nodeValue _
						& " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">New Recipients: " _ 
						& ckCurrentRuleArray(1) & " </font></td></tr>" & vbcrlf	
					Else
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  DisplayActionType(xnRule1.attributes.getNamedItem("ActionType").nodeValue) & " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">Old Arg: " &   xnRule1.attributes.getNamedItem("Arg").nodeValue _
						& " </font></td>" & vbcrlf
						hrmodRuleHtmlReport = hrmodRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">New Arg: " _ 
						& ckCurrentRuleArray(1) & " </font></td></tr>" & vbcrlf	
					End If
				 End if
			End if

		Else
				rrRightReport = 1 
				hrDelRuleHtmlReport = hrDelRuleHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & rnRuleName1 & " </font></td>" & vbcrlf
				hrDelRuleHtmlReport = hrDelRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  DisplayActionType(xnRule1.attributes.getNamedItem("ActionType").nodeValue) _
					& "	" &  xnRule1.attributes.getNamedItem("Arg").nodeValue & " </font></td></tr>" & vbcrlf
				Wscript.echo "Found Deleted Rule: " & rnRuleName1
		End if
	Next
	rem Do forward check of Rule
	For Each dkCurrentRuleKey In ckCurrentRules.keys
		If Not pkPreviousRules.exists(dkCurrentRuleKey) Then
			rrRightReport = 1 
			ckCurrentRuleArray = ckCurrentRules(dkCurrentRuleKey)
			hrnewRuleHtmlReport = hrnewRuleHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">"  _
			& "	" & dkCurrentRuleKey & " </font></td>" & vbcrlf
			hrnewRuleHtmlReport = hrnewRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" & DisplayActionType(ckCurrentRuleArray(0)) & " </font></td>" & vbcrlf
			hrnewRuleHtmlReport = hrnewRuleHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">"  _
			& "	" & ckCurrentRuleArray(1) & " </font></td></tr>" & vbcrlf
			Wscript.echo "Found new Rule : "  & dkCurrentRuleKey 
		End if
	Next
Else
	wscript.echo "No current permissions snap exists taking snap"
	Call TakeSnap
End If

If rrRightReport = 1 Then
	wscript.echo "Writing Report"
	hrHtmlReport = "<html><body>" & vbcrlf
	NewSnapDate = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" 
	hrHtmlReport = hrHtmlReport  & "<p><font size=""4"" face=""Arial Black"" color=""#008000"">Change To Rules or Delgated Folder Permissions for " & unDisplayname & "<Br> for Snaps Taken Between - </font>" & oldSnap  & " and "_
	&  NewSnapDate & "</font></p>" & vbcrlf
	If hrnewHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Added</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrnewHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	If hrmodHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Modified</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrmodHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	If hrDelHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Deleted</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrDelHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	If hrnewRuleHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">New Rule Added</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & hrnewRuleHtmlReport & "</table>"
	End If
	If hrDelRuleHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">Rule Deleted</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & hrDelRuleHtmlReport & "</table>"
	End If
	If hrmodRuleHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">Rule Modified</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & hrmodRuleHtmlReport & "</table>"
	End If
	hrHtmlReport = hrHtmlReport  &  "</body></html>" & vbcrlf
	rfReportFileName = rfReportFileName & Replace(Replace(Replace(NewSnapDate,":",""),",","")," ","")  & ".htm"
	wscript.echo rfReportFileName
	set rfile = fso.opentextfile(rfReportFileName,2,true) 
	rfile.writeline(hrHtmlReport)
End If 

Sub TakeSnap

set wfile = fso.opentextfile(csCurrentSnapFileName,2,true) 
wfile.writeline("<?xml version=""1.0""?>")
wfile.writeline("<SnappedACLS SnapDate=""" & WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" & """>")
Set objSession   = CreateObject("Redemption.RDOSession")
On Error resume next
objSession.LogonExchangeMailbox mbMailboxName, snServerName
if err.number <> 0 Then
	wscript.echo "logon Error"
	wscript.echo err.description
	err.clear
End if
On Error goto 0
set objCuser = objSession.CurrentUser
unDisplayname = objCuser.Name
Set CdoInfoStore = objSession.Stores.DefaultStore
Set CdoFolderRoot = CdoInfoStore.IPMRootFolder
Set FolderACEs = CdoFolderRoot.ACL
For each fldace in FolderACEs
	If  fldace.EntryID = "" Then
		Recpaddress = fldace.Name
		recpName = fldace.Name
	else
		Recpaddress = objSession.GetAddressEntryFromID(fldace.EntryID).Address
		recpName = objSession.GetAddressEntryFromID(fldace.EntryID).Name
	End if
	if cstr(objCuser.address) <> cstr(Recpaddress) Then
		If fwFirstWrite = 0 Then
			wfile.writeline("<Folder Name=""root"">")
			fwFirstWrite = 1
		End if
		wfile.writeline("	<ACE User=""" & Recpaddress & """ Right=""" & DispACERules(fldace) & """ Name = """ & recpName & """></ACE>")
	end if
Next
If fwFirstWrite = 1 then
	wfile.writeline("</Folder>")
End if
Set CdoFolders = CdoFolderRoot.Folders
Set CdoFolder = CdoFolders.GetFirst
do while Not (CdoFolder Is Nothing)
	Set FolderACEs = CdoFolder.ACL
	fwFirstWrite = 0
	For each fldace in FolderACEs
		If  fldace.EntryID = "" Then
			Recpaddress = fldace.Name
			recpName = fldace.Name
		else
			Recpaddress = objSession.GetAddressEntryFromID(fldace.EntryID).Address
			recpName = objSession.GetAddressEntryFromID(fldace.EntryID).Name
		End if
		if cstr(objCuser.address) <> cstr(Recpaddress) Then
			If fwFirstWrite = 0 Then
				wfile.writeline("<Folder Name=""" &  CdoFolder.Name & """>")
				fwFirstWrite = 1
			End if
			wfile.writeline("	<ACE User=""" & Recpaddress & """ Right=""" & DispACERules(fldace) & """ Name = """ & recpName & """></ACE>")
		end if
	Next
	If fwFirstWrite = 1 then
		wfile.writeline("</Folder>")
	End if
	Set CdoFolder = CdoFolders.GetNext
Loop
Set mrMailboxRules = objSession.Stores.DefaultStore.Rules
Wscript.echo "Checking Rules"
fwFirstWrite = 0
bnum = 0
Set dupRules = CreateObject("Scripting.Dictionary")
for Each roRule in mrMailboxRules
	If fwFirstWrite = 0 Then
		wfile.writeline("<Rules>")
		fwFirstWrite = 1
	End If
	agrstr = ""
	acActType = ""
	rname = ""
    set actions = roRule.Actions 
	for i = 1 to actions.count
		acActType = actions(i).ActionType
		if acActType = 6 Or acActType = 8 Or acActType = 7 Then	
		    If acActType = 8 Then
				rname = "Delegate-Forward-Rule-" & bnum
				bnum = bnum + 1
			End if
			for each aoAdressObject In actions(i).Recipients			
				If agrstr = "" then
					agrstr = agrstr & aoAdressObject.Name
				Else 
					agrstr = agrstr & ";" & aoAdressObject.Name
				End if
			next	
		end If
		If acActType = 1 Or acActType = 5 Then
			agrstr = agrstr & actions(i).Folder.Name & " "
		End if
	next
	argstr = fwAdddress
	If roRule.Name = "" And rname = "" Then
			rname = "Blank-" & acActType
	Else
		If rname = "" Then
			rname = Replace(Replace(roRule.Name,"<"," "),">"," ")
		End if
	End If
	If dupRules.exists(rname) Then 
		wscript.echo "Duplicate in Rules Founds #########################"
	Else
		dupRules.add rname,1
		wfile.writeLine("	<Rule Name =""" & rname & """ ActionType=""" & acActType & """ Arg=""" & agrstr & """></Rule>")
	End if
Next
If fwFirstWrite = 1 then
		wfile.writeline("</Rules>")
End if

if Not objSession Is Nothing Then objSession.Logoff 
set objSession = Nothing
Set mrMailboxRules = Nothing
wfile.writeline("</SnappedACLS>")
wscript.echo "New Snap Taken"

End Sub

Function DispACERules(DisptmpACE)

Select Case DisptmpACE.Rights

        Case ROLE_NONE, 0  ' Checking in case the role has not been set on that entry.
                DispACERules = "None"
        Case 1024  ' Check value since ROLE_NONE is incorrect
                DispACERules = "None"
        Case ROLE_AUTHOR
                DispACERules = "Author"
        Case 1051  ' Check value since ROLE_AUTHOR is incorrect
                DispACERules = "Author"
        Case ROLE_CONTRIBUTOR
                DispACERules = "Contributor"
        Case 1026  ' Check value since ROLE_CONTRIBUTOR is incorrect
                DispACERules = "Contributor"
        Case 1147  ' Check value since ROLE_EDITOR is incorrect
                DispACERules = "Editor"
        Case ROLE_NONEDITING_AUTHOR
                DispACERules = "Nonediting Author"
        Case 1043  ' Check value since ROLE_NONEDITING AUTHOR is incorrect
                DispACERules = "Nonediting Author"
        Case 2043  ' Check value since ROLE_OWNER is incorrect
                DispACERules = "Owner"
        Case ROLE_PUBLISH_AUTHOR
                DispACERules = "Publishing Author"
        Case 1179  ' Check value since ROLE_PUBLISHING_AUTHOR is incorrect
                DispACERules = "Publishing Author"
        Case 1275  ' Check value since ROLE_PUBLISH_EDITOR is incorrect
                DispACERules = "Publishing Editor"
        Case ROLE_REVIEWER
                DispACERules = "Reviewer"
        Case 1025  ' Check value since ROLE_REVIEWER is incorrect
                DispACERules = "Reviewer"
        Case Else
                DispACERules = "Custom"
End Select

End Function

Function DisplayActionType(acActionType)
	Select Case acActionType
		Case 1 DisplayActionType = "Move-Rule"
		Case 2 DisplayActionType = "Assign Cateogry"
		Case 3 DisplayActionType = "Delete-Message"
		Case 4 DisplayActionType = "Delete-Permanently"
		Case 5 DisplayActionType = "Copy-Rule"
		Case 6 DisplayActionType = "Forward-Rule"
		Case 7 DisplayActionType = "Forward-As-Attachment"
		Case 8 DisplayActionType = "Delegate-Forward-Rule"
		Case 9 DisplayActionType = "ServerReply "
		Case 11 DisplayActionType = "Mark-Defer"
		Case 15 DisplayActionType = "Importance"
		Case 16 DisplayActionType = "Sensitivity"
		Case 19 DisplayActionType = "Mark-Read"
		Case 28 DisplayActionType = "Defer"
		Case 1024 DisplayActionType = "Bounce-Message"
		Case 1025 DisplayActionType = "Tag"
		Case Else DisplayActionType = "Unknown"
	End Select
End Function 



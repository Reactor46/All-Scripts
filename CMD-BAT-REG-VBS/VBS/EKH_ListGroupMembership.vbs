'****************************************************************************************************
'*** ListAllGroupsAndMembers.vbs - Jim Turner / modified by Andreas Iwanowski
'*** Enumerates all Domain groups
'*** Lists Nested Groups
'*** Lists users whos Primary Group is the group being enumerated
'*** Hilites Groups in Red and Reoccurring Groups in Purple.
'*** Hilites Disabled Accounts in Blue.
'*** Takes care of possible endless loops should one group contain another group that contains it.
'*** Sorts and indents by Group
'*** Provides Summary list with Links to Groups.
'*** Provides No Members list with Links to Groups without Members (if any exists).
'*** Creates Red worksheet tabs for Groups with No Members.
'*** Requires Excel and user must have Domain Admin level access
'*** Places sorting string in column IV (so that it's out of the way)  (IV is column 256)
'****************************************************************************************************
Public Row,Col,savegroup,HasMembers
Const ADS_SCOPE_SUBTREE = 2
Const xlAscending = 1
Const xlHeader = 1
On Error Resume Next

Const xlMinimized = -4140 '(&HFFFFEFD4)
Const xlNormal = -4143    '(&HFFFFEFD1)

strMessage = "Spreadsheet will open when process is Done"
strScriptName = "List All Groups And Members"
'CreateObject("WScript.Shell").Popup strMessage,3,strScriptName,vbInformation

Row = 1 : Col = 1 : GroupCount = 0 : NoMembers = 0 : YesMembers = 0
StartTime = Now()
Set XL = CreateObject("Excel.Application")
If Err.Number <> 0 Then
	WScript.Echo "Microsoft Excel does not appear to be installed on this system."
	WScript.Quit
End If

XL.Workbooks.Add
XL.WindowState = xlMinimized

'*** add code to hilite disabled accounts
Set DisabledAcct = CreateObject("Scripting.Dictionary")

'*** All domains
Const nNumDomains = 2
Dim arDNCs()
ReDim arDNCs(nNumDomains)

arDNCs(0)           = "contoso.com/DC=contoso,DC=com"
arDNCs(1)           = "subA.contoso.com/DC=subA,DC=contoso,DC=com"
arDNCs(2)           = "subB.contoso.com/DC=subB,DC=contoso,DC=com"




If Err.Number <> 0 Then
	WScript.Echo Err.Number
	WScript.Echo "Could not get the default naming context; is this system a domain member?"
	XL.Quit
	WScript.Quit
End If


Set objConn = CreateObject("ADODB.Connection")
Set objComm = CreateObject("ADODB.Command")
objConn.Provider = "ADsDSOObject"
objConn.Open "Active Directory Provider"
Set objComm.ActiveConnection = objConn
objComm.Properties("Page Size") = 1000

LDAPFilter = "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"

Dim arObjRec()
ReDim arObjRec(nNumDomains)
i = 0
For Each DNC In arDNCs
	strQuery = "<LDAP://" & DNC & ">;" & LDAPFilter & ";adspath,samaccountname;subtree"
	objComm.CommandText = strQuery
	Set arObjRec(i) = objComm.Execute
	i = i + 1
Next

Set objRec = MergeRecordSets(arObjRec)

objRec.MoveFirst

Do Until objRec.EOF
 strAccName = DomainNameFromAdsPath(objRec.Fields("adspath").Value) & "\" & objRec.Fields("samaccountname").Value
 DisabledAcct.Add strAccName,strAccName
 objRec.MoveNext
Loop
objConn.Close

Set HasMemberList = CreateObject("Scripting.Dictionary")
Set NoMemberList = CreateObject("Scripting.Dictionary")

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Sort On") = "CN"
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

Dim arObjRecordsets()
ReDim arObjRecordsets(nNumDomains)
j = 0
For Each DNC in arDNCs
    objCommand.CommandText = "SELECT ADsPath,cn,primaryGroupToken FROM 'LDAP://" & DNC & "' WHERE objectclass = 'group'"
	Set arObjRecordsets(j) = objCommand.Execute
	j = j + 1
Next

Set objRecordSet = MergeRecordSets(arObjRecordsets)

objRecordSet.MoveFirst

'*** Step thru all groups
Do Until objRecordSet.EOF
 HasMembers = False
 Row = 1
 Set dictObj = CreateObject("Scripting.Dictionary")
 strPath = objRecordSet.Fields("ADsPath").Value
 Set GrpObj = GetObject(strPath)
 strGroupFullName = GetGroupMemberFullName(GrpObj)

 shname = Replace(strGroupFullName,"\","_")
 '*** Excel sheetname limit 31 characters
 If Len(shname) > 31 Then
  shname = Left(shname,31)
 End If
 XL.Sheets.Add.Name = shname
 XL.Cells(Row,1).Value = strGroupFullName
 XL.Cells(Row,256).Value = " " & strGroupFullName
 Row = Row + 1
 XL.Rows("1:1").Font.Color = RGB(0,128,0)  'Green Group Heading
 XL.Rows("1:1").Font.Bold  = True 'Bold Heading
 
 savegroup = GrpObj.name
 sortgroup = savegroup
 GroupCount = GroupCount + 1

 '*** Add Source Group to Group Dictionary list
 dictObj.Add strGroupFullName,strGroupFullName

 '*** Get Group Members
 GetGroupMembers(GrpObj)

 If Not HasMembers Then
  XL.Activesheet.Tab.ColorIndex = 3
  NoMembers = NoMembers + 1
  NoMemberList.Add XL.Cells(1,1).Value,XL.Activesheet.name
 Else
  '*** Sort Spreadsheet
  'XL.Range("A1").Select
  'XL.Cells.Select
  'XL.Selection.Sort XL.Range("IV1"),xlAscending,,,,,,xlHeader,1,False
  'XL.Range("A1").Select
  'YesMembers = YesMembers + 1
  '*** Key contains full group name, Item contains truncated worksheet name
  HasMemberList.Add XL.Cells(1,1).Value,XL.Activesheet.name
 End If



'*** For Printing - Must Hide Sort column IV
XL.Columns("IV:IV").Select
XL.Selection.EntireColumn.Hidden = True
'*** Set Footer to contain Worksheet name and Page number
XL.ActiveSheet.PageSetup.PrintArea = ""
With XL.ActiveSheet.PageSetup
 .CenterFooter = "&A"
 .RightFooter = "Page &P"
End With
XL.Range("A1").Select



  objRecordSet.MoveNext
Loop

objConnection.Close

If NoMembers > 0 Then
 '*** Create Groups with No Members Worksheet
 XL.Sheets.Add.Name = "GroupsWithNoMembers"
 XL.Activesheet.Tab.ColorIndex = 3
 Row = 1
 XL.Cells(Row,1).Value = NoMembers & " Groups with No Members:"
 Row = Row + 1
 GroupKeys2 = NoMemberList.Items()
 GroupKey2a = NoMemberList.Keys()
 For i = 0 To NoMemberList.Count - 1
  '*** Write full groupname but provide truncated worksheet name as link
  '*** NoMemberList key contains Full Group Name
  '*** NoMemberList Item contains worksheet name and will be truncated to 31 characters
  XL.Cells(Row,1).Value = GroupKey2a(i)
  Linkname = GroupKeys2(i)
  If Len(Linkname) > 31 Then
   HLink = "'" & Left(Linkname,31) &"'!a1"
  Else
   HLink = "'" & GroupKeys2(i) &"'!a1"
  End If
  XL.Range("A" & Row).Select
  XL.ActiveSheet.Hyperlinks.Add XL.Selection,"",HLink,"Link to " & GroupKeys2(i)
  Row = Row + 1
 Next



'*** Set Footer to contain Worksheet name and Page number
XL.ActiveSheet.PageSetup.PrintArea = ""
With XL.ActiveSheet.PageSetup
 .CenterFooter = "&A"
 .RightFooter = "Page &P"
End With



 XL.Range("A1").Select
 XL.Cells.EntireColumn.AutoFit
End If

'*** Create Summary Worksheet
XL.Sheets.Add.Name = "SummaryOfGroups"
XL.Activesheet.Tab.ColorIndex = 4
Row = 1
XL.Cells(Row,1).Value = "Total Group Count"
XL.Cells(Row,2).Value = GroupCount

If NoMembers > 0 Then
 '*** Add Groups with No Members summary info if applicable
 Row = Row + 1
 XL.Cells(Row,1).Value = "Groups No Members"
 XL.Cells(Row,2).Value = NoMembers
 XL.Cells(Row,3).Value = "See GroupsWithNoMembers worksheet tab"
 XL.Range("C" & Row).Select
 XL.ActiveSheet.Hyperlinks.Add XL.Selection,"","'" & "GroupsWithNoMembers" &"'!a1","Link to " & "GroupsWithNoMembers"
End If

Row = Row + 2

'*** Create Hyperlinks to Groups with members worksheets
XL.Cells(Row,1).Value = YesMembers & " Groups with Members:"
Row = Row + 1
GroupKeys1 = HasMemberList.Items()
GroupKey1a = HasMemberList.Keys()
For i = 0 To HasMemberList.Count - 1
 '*** Write full groupname but provide truncated worksheet name as link
 '*** HasMemberList key contains Full Group Name
 '*** HasMemberList Item contains worksheet name and will be truncated to 31 characters
 XL.Cells(Row,1).Value = GroupKey1a(i)
 Linkname = GroupKeys1(i)
 If Len(Linkname) > 31 Then
  HLink = "'" & Left(Linkname,31) &"'!a1"
 Else
  HLink = "'" & GroupKeys1(i) &"'!a1"
 End If
 XL.Range("A" & Row).Select
 XL.ActiveSheet.Hyperlinks.Add XL.Selection,"",HLink,"Link to " & GroupKeys1(i)
 Row = Row + 1
Next

EndTime = Now()
XL.Cells(Row + 1,1).Value = "StartTime"
XL.Cells(Row + 1,2).Value = StartTime
XL.Cells(Row + 2,1).Value = "EndTime"
XL.Cells(Row + 2,2).Value = EndTime



'*** Set Footer to contain Worksheet name and Page number
XL.ActiveSheet.PageSetup.PrintArea = ""
With XL.ActiveSheet.PageSetup
 .CenterFooter = "&A"
 .RightFooter = "Page &P"
End With



XL.Range("A1").Select
XL.Cells.EntireColumn.AutoFit

XL.WindowState = xlNormal
XL.Visible = TRUE

strMessage = "List All Groups And Members process is Complete"
strScriptName = "List All Groups And Members"
CreateObject("WScript.Shell").Popup strMessage,45,strScriptName,vbInformation

''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
Sub GetGroupMembers(Grp)
 On Error Resume Next
 '*** Call Out C
 '*** Some Admins tend to make their Primary Group Domain Admins
 '*** Those will not show up when querying group membership
 '*** Need to look for users that have this groups primaryGroupToken
 Grp.GetInfoEx Array("primaryGroupToken"),0
 tokName = Grp.Get("primaryGroupToken")
 LDAPFilter = "(primaryGroupID=" & tokName & ")"

 strQuery = "<LDAP://" & DNCFromGroupObject(Grp) & ">;" & LDAPFilter & ";adspath,samaccountname,cn;subtree"
 objCommand.CommandText = strQuery
 Set objRecordset2 = objCommand.Execute

 objRecordset2.MoveFirst

 If objRecordset2.RecordCount > 0 Then
     Do Until objRecordset2.EOF
      HasMembers = True
      strFullName = DomainNameFromAdsPath(objRecordset2.Fields("adspath").Value) & "\" & objRecordset2.Fields("samaccountname").Value
      XL.Cells(Row,col).Value = strFullName & " (" & objRecordset2.Fields("cn").Value & ")"
      XL.Cells(Row,256).Value = sortgroup & " " & objRecordset2.Fields("cn").Value

      USam = ""
      USam = strFullName

      If DisabledAcct.Exists(USam) Then
       XL.Rows(Row & ":" & Row).Font.Color = RGB(0,0,255)
       XL.Cells(Row,col).Value = XL.Cells(Row,col).Value + " (Disabled Account)"
      End If

      Row = Row + 1
      objRecordset2.MoveNext
     Loop
 End If
 objRecordset2.Close
 '*** End Call Out C

 '***Call Out D
 For Each memobj In Grp.Members

  strFullName = GetGroupMemberFullName(memobj)

  If Lcase(memobj.Class) = "group" Then
   HasMembers = True
   '*** Add to dictionary if it does not already exist 
   If Not dictObj.Exists(strFullName) Then     
    dictObj.Add strFullName,strFullName
    XL.Rows(Row & ":" & Row).Font.Color = RGB(255,0,0) 'make groupname red
    XL.Rows(Row & ":" & Row).Font.Bold = True 'groupnames in bold
    XL.Cells(Row,Col).Value =  strFullName  & " (" & memobj.cn & ")"
    '*** Since Group changed, concat groupname to sortgroup
    '*** asterisk is used as a separator
    sortgroup = sortgroup + "*" + memobj.name
    XL.Cells(Row,256).Value =  sortgroup
    Row = Row + 1
    '*** increase indent for new group
    Col = Col + 1
    '*** Store new groupname to savegroup
    savegroup = memobj.name
    '*** recurse subgroups
    GetGroupMembers(memobj)
   Else
    '*** This group has already been enumerated.  Make reoccurring groupname purple
    XL.Rows(Row & ":" & Row).Font.Color = RGB(255,0,255)
    XL.Cells(Row,col).Value =  strFullName & " (Reoccurring Group)"
    '*** Since Group changed concat groupname to sortgroup
    '*** asterisk is used as a separator
    sortgroup = sortgroup + "*" + memobj.name
    XL.Cells(Row,256).Value =  sortgroup
    Row = Row + 1
    '*** Since not enumerating group - remove reoccurring groupname from sortgroup
    sortgroup = Left(sortgroup,instrrev(sortgroup,"*")-1)
   End If
  Else
   HasMembers = True
   '*** ... not a group
   XL.Cells(Row,col).Value =  strFullName  & " (" & memobj.cn & ")"
   XL.Cells(Row,256).Value =  sortgroup & " " &  memobj.cn

   USam = ""
   USam = strFullName
   If DisabledAcct.Exists(USam) Then
    XL.Rows(Row & ":" & Row).Font.Color = RGB(0,0,255)
    XL.Cells(Row,col).Value = XL.Cells(Row,col).Value + " (Disabled Account)"
   End If

   Row = Row + 1
  End If
  '*** Remove indent if group is no longer the same
  If savegroup <> grp.name Then
   Col = Col - 1
  '*** Store current groupname to savegroup
   savegroup = grp.name
  '*** Remove last group from sortgroup string
   sortgroup = Left(sortgroup,instrrev(sortgroup,"*")-1)
  End If
 Next
 '*** End Call Out D
 Set memobj = Nothing
 Set objRecordset2 = Nothing
End Sub



'*** EKH: Merge recordsets (e.g. multiple domain results)
Function MergeRecordSets(arrRecordsets)
  
	Dim x, y, objCurrentRS
	Dim objMergedRecordSet, objField, blnExists
	Set objMergedRecordSet = CreateObject("ADODB.Recordset")
	For x=0 To UBound(arrRecordsets)
		Set objCurrentRS = arrRecordsets(x)

	    For Each objField In objCurrentRS.Fields
		    blnExists = False
		    For y=0 To objMergedRecordSet.Fields.Count-1
			    If LCase(objMergedRecordSet.Fields(y).Name) = Lcase(objField.Name) Then
				    blnExists = True : Exit For
			    End If
		    Next
		    If Not(blnExists) Then
			    objMergedRecordSet.Fields.Append objField.Name, objField.Type, objField.DefinedSize
			    'objMergedRecordSet.Fields(objMergedRecordset.Fiel    ds.Count-1).Attributes = 32 'adFldIsNullable
		    End If
	    Next
	Next
	
	objMergedRecordSet.Open
	
	For x=0 To UBound(arrRecordsets)
		Set objCurrentRS = arrRecordsets(x)
        If objCurrentRS.RecordCount > 0 Then
	        Do Until objCurrentRS.EOF
		        objMergedRecordSet.AddNew
		        For Each objField In objCurrentRS.Fields
			        If Not(IsNull(objField.Value)) Then
				        objMergedRecordSet.Fields(objField.Name).Value = objField.Value
				    End If
			    Next
		    objCurrentRS.MoveNext
            Loop

        End If
    Next


	
	If Not objMergedRecordSet.EOF Then
        objMergedRecordSet.MoveFirst
    End If
	Set MergeRecordSets = objMergedRecordSet


End Function

'*** EKH: Get Full group name
Function GetGroupMemberFullName(memobj)

   strDomain = DomainNameFromAdsPath(memobj.AdsPath)
   
   strFullName = strDomain + "\" + memobj.samaccountname
   GetGroupMemberFullName = strFullName
End Function

Function DomainNameFromAdsPath(adspath)
   aElements = Split(adspath,",DC=",-1,vbTextCompare)
   strDomain = ""

   '''''''''''''''''''''''''
   ' Enable this for domain FQDN
   'For i = 1 to UBound(aElements)
   ' strDomain = strDomain + "." + aElements(i)
   'Next
   'strDomain = Mid(strDomain, 2)
   '''''''''''''''''''''''''
   
   '''''''''''''''''''''''''
   ' Enable this for SL domain name
   strDomain = UCase(aElements(1))
   '''''''''''''''''''''''''

   DomainNameFromAdsPath = strDomain


End Function

Function DNCFromGroupObject(memobj)
   strAds = memobj.AdsPath
   aElements = Split(memobj.Adspath,",DC=",-1,vbTextCompare)

   strDomain = ""
   strDNC = ""

   For i = 1 to UBound(aElements)
     strDomain = strDomain + "." + aElements(i)
     strDNC = strDNC + ",DC=" + aElements(i)
   Next

   strDomain = Mid(strDomain, 2)
   strDNC = Mid(strDNC, 2)

   DNCFromGroupObject = strDomain + "/" + strDNC

End Function
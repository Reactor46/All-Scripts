'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' http://sites.google.com/site/assafmiron
' Date : 18/09/2010
' Get PST File Path and Size.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
'======================================================================================================================
' Consts
'======================================================================================================================
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
Const tmpSeperator = ";"
'File Consts
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
'DB Consts
Const adPersistXML = 1
Const adVarChar = 200
Const MaxCharacters = 500
Const adFldIsNullable = 32
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
' File Size Consts
Const FileExceed = 300000
' Mail Consts
Const strSubject = "PST File Exceeded File Size"
'======================================================================================================================
' Set Objects
'======================================================================================================================
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set DataList = CreateObject("ADOR.Recordset")
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")

Sub SendMail(ToAddress, CCAddress, BCCAddress, MessageSubject, MessageBody, AttachmentPath)
	Const olMail = 0
	Dim objFSO, objOutlook, objMail
	' Create The File System Object - To check if the File Exsists
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	' Create the Outlook Application Object
	Set objOutlook = CreateObject("Outlook.Application")
	Set objMail = objOutlook.CreateItem(olMail)
	' Set the Message Properties
	With objMail
        .To = ToAddress ' Send To Address
        .CC = CCAddress ' Add a CC Address
        .BCC = BCCAddress ' Add a BCC Address
        .Subject = MessageSubject ' Set the Message Subject
        .Body = MessageBody & vbCrLf ' Set the Message Body
        ' Check that there is an Attachment
        If Not AttachmentPath = "" Then
        	' Add an Attachment if it Exsits
        	If objFSO.FileExists(AttachmentPath) Then
        		.Attachments.Add AttachmentPath 
        	Else
        		MsgBox "Attachment Path Does Not Exists"
        	End If
        End If
        .Display ' Show the Message
        '.Send ' Send the Message
    End With
    ' Clean Up
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub

Function FindPST(strComputer)
'======================================================================================================================
' Sub recevies computer name and checks its registry for the DefaultProfile Key to retrive the default Outlook Profile name.
' Then Checks the 001f6700 Key in each profile for the path of all the PST files that are connected to that Profile.
' Makes use of the RegBintoString Function to translate the Reg Key value.
'======================================================================================================================
	Dim arrPstPath()
	Dim PstPath
	Dim i : i = 0
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
	 strComputer & "\root\default:StdRegProv")
	'======================================================================================================================
	' Connect to computer Registry and retreve the Parent value
	'======================================================================================================================
	'Get The Default Profile Name
	strOlkProfileKey = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
	strValueName = "DefaultProfile"
	oReg.GetStringValue HKEY_CURRENT_USER,strOlkProfileKey,strValueName,strValue
	DefProf = strValue
	Set objPSTFile = objFSO.CreateTextFile("C:\tmpPst.txt")
	objPSTFile.WriteLine "PST"
	' Get The Pst Path in the Default Profile
	strKeyPath = strOlkProfileKey & "\" & DefProf 
	strValueName = "001f6700"
	oReg.EnumKey HKEY_CURRENT_USER,strKeyPath, arrSubKeys
	For Each strKey in arrSubKeys
		strSubKey = strKeyPath & "\" & strKey
		oReg.GetBinaryValue HKEY_CURRENT_USER,strSubKey,strValueName,arrValue
		If Not IsNull(arrValue) Then
			PSTPath = RegBintoString(arrValue)
		End If
		If Not PSTPath = "" Then
		'Write the Pst Paths to a tmp File
			objPSTFile.WriteLine PSTPath
		End If
	Next
	objPSTFile.Close
	strPathtoTextFile = "C:\"
	'======================================================================================================================
	' Tally the Temp File
	'======================================================================================================================
	objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	    "Data Source=" & strPathtoTextFile & ";" & _
	        "Extended Properties=""text;HDR=Yes;FMT=Delimited"""
	strFile = "tmpPst.txt"
	objRecordset.Open "Select Distinct(PST) FROM " & strFile & _
	    " GROUP BY PST", objConnection, adOpenStatic, adLockOptimistic, adCmdText
	ReDim arrPstPath(objrecordset.RecordCount-1)
	Do Until objRecordset.EOF
	    arrPstPath(i) = objRecordset.Fields.Item("PST")
	    objRecordset.MoveNext
	    i = i + 1
	Loop
	objConnection.Close
	objFSO.DeleteFile("C:\tmpPst.txt")
	FindPST = arrPstPath
End Function

Function RegBintoString(arrValue)
'======================================================================================================================
' Converts Binary Registry Values to String
'======================================================================================================================
	For i = 0 to UBound(arrValue) step 2
		tmpText = tmpText & Chr(arrValue(i))
	Next
	RegBintoString = tmpText
End Function

Private Function fFormatNum(num, DropDecimal)
'======================================================================================================================
' Format a Number by size and adds the apropriate bytes description - MB,KB,GB...
'======================================================================================================================
	Dim bytes
	Dim lngSize
	If IsNumeric(num) Then
		If Len(num) < 5 Then
			lngSize = FormatNumber((num /1024), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Gb"
		Elseif Len(num) < 7 Then
			lngSize = FormatNumber((num / 1024), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Kb"
		Elseif Len(num) < 10 Then
			lngSize = FormatNumber((num / 1024 ^ 2), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Mb"
		Elseif Len(num) < 13 Then
			lngSize = FormatNumber((num / 1024 ^ 3), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Gb"
		Elseif Len(num) < 16 Then
			lngSize = FormatNumber((num / 1024 ^ 4), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Tb"
		Elseif Len(num) >= 16 Then
			lngSize = FormatNumber((num / 1024 ^ 5), 2, vbTrue, vbTrue, vbUseDefault)
			bytes = " Pb"
		End If
		If DropDecimal Or Right(lngSize, 2) = "00" Then
			fFormatNum = CStr(Round(lngSize, 0)) & bytes
		Else
			fFormatNum = CStr(lngSize) & bytes
		End If
	Else
		fFormatNum = num
	End If
End Function

Function GetFileSize(strFilePath)
' Getting File Size
	Set objFile = objFSO.GetFile(strFilePath)
	GetFileSize = objFile.Size
End Function

' Get the PST File Paths Array
arrPSTPath = FindPST(".")
' Define the Send to EMail Address
strSendTo = "assaf.miron@gmail.com"
For Each strPstPath In arrPSTPath
	' See the PST Path
	WScript.Echo strPSTPath
	' Get the File Size
	iFileSize = GetFileSize(StrPSTPath)
	' Echo the File Size Formatted
	WScript.Echo fFormatNum(iFileSize,False)
	' Check the File Size, if it Exceeds the File Size Defined
	If iFileSize > FileExceed Then
		WScript.Echo "The File Exceeded the Normal File Size!" & vbNewLine & "Sending Mail to " & strSendTo
		' Set the Message Body
		strBody = "Dear client, " & vbNewLine & "You PST File Exceeded the Normal File Size Defined." & vbNewLine 
		strBody = strBody & "Your PST File  is " & fFormatNum(iFileSize,False)		
		' Send the Mail
		SendMail strSendTo,"","" strSubject, strBody, ""
	End If
Next

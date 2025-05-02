'***************************************************************************************************
'*                                                                                                 *
'*                                       OutlookSigGenerator.vbs                                   *
'*                                       Written By: ITCowboy42                                    *
'*                                                                                                 *
'*                                        Version 2.1.3                                            *
'*                                       Original Date: 4/5/2005                                   *
'*                                        Revised:  4/15/2011                                      *
'*  This script gets user info from Active Directory and uses that information to create Outlook   *
'*   Signature in Text, Html, and RTF fashion and saves it to the appropriate files for user.      *
'*                                                                                                 *
'*   Read through and test code carefully as many fields will need changes, and specific formatting*
'*                           to look professional. This is just a starting area.                   *
'*                                                                                                 *
'***************************************************************************************************

Option Explicit

Dim qQuery, objFSO, objFolder
Dim objSysInfo
Dim objuser
Dim strDisclaimer
Dim strProfilePath
Dim oshell
Set objSysInfo = CreateObject("ADSystemInfo")
'objSysInfo.RefreshSchemaCache
qQuery = "LDAP://" & objSysInfo.Username

Dim FullName
Dim EMail
Dim Title
Dim PhoneNumber
Dim FaxNumber
Dim OfficeLocation
Dim web_address
Dim CompanyName
Dim strIPPhone
Dim BlnVoice
Dim StreetAddress
Dim Town
Dim State
Dim ZipCode

CompanyName = "Company Name Here"
'strDisclaimer = "This email (and or documents accompanying it) may contain confidential" &_
'	" information belonging to a sender, which is protected by the consultant-client" &_
'	" privilege.  The information is intended only for the use of the individual or" &_
'	" entity named above.  If you are not the intended recipient, you are hereby" &_
'	" notified that any disclosure, copying, distribution, or the taking of any action" &_
'	" in reliance on the contents of this information is strictly prohibited.  If you" &_
'	" received this email in error, please notify the sender immediately and delete" &_
'	" the documents from your computer."
strDisclaimer = "Please note: This email (and or documents accompanying it)" &_ 
    " is directed in confidence solely to the person(s) directed" &_
	" above. The contents of this email are confidential. We are aware that" &_
	" occasional transmission errors occur in routing email messages. If you have" &_
	" received this message in error, please notify us immediately, and destroy the" &_
	" email without making a copy. Thank you."
	
' Asigns the user's info to variables.
'==========================================================================
set oShell = WScript.CreateObject("WScript.Shell") 
strProfilePath = oShell.ExpandEnvironmentStrings("%USERPROFILE%") 

Set objuser = GetObject(qQuery)
FullName = objuser.displayname
EMail = objuser.mail
Title = objuser.title
PhoneNumber = objuser.telephoneNumber
FaxNumber = objuser.facsimileTelephoneNumber
OfficeLocation = objuser.physicalDeliveryOfficeName
strIPPhone = objuser.ipPhone 'We used the ip phone category from AD to store individual mailbox numbers

' This section looks for the signatures folder and if it does not exist, creates it
'==========================================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strProfilePath & "\Application Data\Microsoft\Signatures\") Then
'   Set objFolder = objFSO.GetFolder(strUserProfile & "\Application Data\Microsoft\Signatures")
'    wscript.Echo "Folder Exists"
Else
    Set objFolder = objFSO.CreateFolder(strProfilePath & "\Application Data\Microsoft\Signatures")

'    Wscript.Echo "Folder created"
End If

' This section creates the signature files names and locations.
'==========================================================================
Dim FolderLocation
Dim LogonID
Dim TexTFileString
Dim RTFFileString
Dim HTMFileString

web_address = "www.yourcompany.com" 'Replace with your web address


LogonID = objuser.sAMAccountName
'FolderLocation = "C:\Documents and Settings\" & LogonID & "\Application Data\Microsoft\Signatures\"
FolderLocation = strProfilePath & "\Application Data\Microsoft\Signatures\"


TexTFileString = FolderLocation & "Default.txt"
RTFFileString = FolderLocation & "Default.rtf"
HTMFileString = FolderLocation & "Default.htm"


' This section figures out where the user works. Uses the location field from AD
'==========================================================================
'Dim StreetAddress
'Dim Town
'Dim State
'Dim ZipCode


'Edit this section as needed.
BlnVoice = 0

If OfficeLocation = "Location 1" Then
PhoneNumber = "(455) 455-3399"
FaxNumber = "(455) 455-3857"
StreetAddress = "Address Line 1"
Town = "Your Town"
State = "YourState"
ZipCode = "zip"
End If

If OfficeLocation = "Location 2" Then
PhoneNumber = "(455) 455-5988"
FaxNumber = "(455) 455-4361"
StreetAddress = "Your 2nd Address"
Town = "Town"
State = "YourState"
ZipCode = "Zip"
BlnVoice =1
End IF

'If OfficeLocation = "Location 3" Then
'StreetAddress = ""
'Town = ""
'State = ""
'ZipCode = ""
'End IF

If OfficeLocation = "" Then
Wscript.echo "Your administrator must enter your location for this script to run. Please advise your IT department."
Wscript.quit
End if



' Thes next 3 sections builds the signature files
'==========================================================================
Dim objFile
Dim aQuote
aQuote = chr(34)

' This section builds the text file version
'==========================================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(TexTFileString)
objFile.Close
Set objFile = objFSO.OpenTextFile(TexTFileString, 2)  
objfile.write CompanyName & vbCrLf & vbCrLf
objfile.write FullName & vbCrLf
objfile.write Title & vbCrLf & vbCrLf
' Replace Company-name with your company name
'objfile.write StreetAddress & vbCrLf
'objfile.write Town & " " & State & " " & ZipCode & vbCrLf
objfile.write "Office   : " & Phonenumber & vbCrLf
If BlnVoice = 0 then
objfile.write "VoiceMail: " & strIPPhone & vbCrLf
End if
objfile.write "Fax      : " & FaxNumber & vbCrLf
objfile.write " " & vbCrLf
objfile.write StreetAddress & vbCrLf
objfile.write Town & ", " & State & " " & ZipCode & vbCrLf
objfile.write " " & vbCrLf
objfile.write web_address & vbCrLf

objFile.Close

' This section builds the HTML file version
'==========================================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(HTMFileString)
objFile.Close
Set objFile = objFSO.OpenTextFile(HTMFileString, 2)  

objfile.write "<!DOCTYPE HTML PUBLIC " & aQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & aQuote & ">" & vbCrLf
objfile.write "<HTML><HEAD><TITLE>Microsoft Office Outlook Signature</TITLE>" & vbCrLf
objfile.write "<META http-equiv=Content-Type content=" & aQuote & "text/html; charset=windows-1252" & aQuote & ">" & vbCrLf
objfile.write "<META content=" & aQuote & "MSHTML 6.00.3790.186" & aQuote & " name=GENERATOR></HEAD>" & vbCrLf
objfile.write "<BODY>" & vbCrLf
objfile.write "<DIV align=left><FONT face=Verdana color=#000000 size=2><STRONG>" & CompanyName &"</STRONG></FONT><BR><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2>" & FullName & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2>" & Title & "</FONT><BR><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2>Office   : " & PhoneNumber & "</FONT><BR>" & vbCrLf
If BlnVoice = 0 then
objfile.write "<FONT face=Verdana color=#000000 size=2>VoiceMail: " & strIPPhone & "</FONT><BR>" & vbCrLf
End if
objfile.write "<FONT face=Verdana color=#000000 size=2>Fax      : " & FaxNumber & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2> " & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2>" & StreetAddress & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2>" & Town & ", " & State & " " & ZipCode & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=2> " & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana size=2>" & "<A href=" & aQuote & web_address & aQuote & ">" & web_address & "</A></FONT><BR><BR><BR>" & vbCrLf
objfile.write "<FONT face=Verdana color=#000000 size=1>" & strDisclaimer & "</FONT><BR>" & vbCrLf
objfile.write "<FONT face=Verdana size=2></FONT>&nbsp;</DIV></BODY></HTML>" & vbCrLf
objFile.Close

' This section builds the RTF file version
'==========================================================================
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile(RTFFileString)
objFile.Close
Set objFile = objFSO.OpenTextFile(RTFFileString, 2)  

objfile.write "{\rtf1\ansi\ansicpg1252\fromhtml1 \deff0{\fonttbl" & vbCrLf
objfile.write "{\f0\ Verdana;}}" & vbCrLf
'objfile.write "{\f1\fmodern Courier New;}" & vbCrLf
'objfile.write "{\f2\fnil\fcharset2 Symbol;}" & vbCrLf
'objfile.write "{\f3\fmodern\fcharset0 Courier New;}" & vbCrLf
'objfile.write "{\f4\fswiss Century Schoolbook;}" & vbCrLf
'objfile.write "{\f5\fswiss Arial;}}" & vbCrLf
'objfile.write "{\colortbl\red0\green0\blue0;\red0\green0\blue0;\red0\green0\blue0;}" & vbCrLf
objfile.write "\uc1\pard\plain\deftab360 \f0\fs20 " & vbCrLf
objfile.write "{\*\htmltag243 <!DOCTYPE HTML PUBLIC " & aQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & aQuote & ">}" & vbCrLf
objfile.write "{\*\htmltag3 \par }" & vbCrLf
objfile.write "{\*\htmltag19 <HTML>}" & vbCrLf
objfile.write "{\*\htmltag34 <HEAD>}" & vbCrLf
objfile.write "{\*\htmltag177 <TITLE>}" & vbCrLf
objfile.write "{\*\htmltag241 Microsoft Office Outlook Signature}" & vbCrLf
objfile.write "{\*\htmltag185 </TITLE>}" & vbCrLf
objfile.write "{\*\htmltag1 \par }" & vbCrLf
objfile.write "{\*\htmltag1 \par }" & vbCrLf
objfile.write "{\*\htmltag161 <META content=" & aQuote & "MSHTML 6.00.3790.186" & aQuote & " name=GENERATOR>}" & vbCrLf
objfile.write "{\*\htmltag41 </HEAD>}" & vbCrLf
objfile.write "{\*\htmltag2 \par }" & vbCrLf
objfile.write "{\*\htmltag50 <BODY>}" & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag84 <STRONG>}\htmlrtf {\b \htmlrtf0 " & CompanyName & vbCrLf
objfile.write "{\*\htmltag92 </STRONG>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 "& Fullname & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 \par size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 " & Title & vbCrLf
'objfile.write "{\*\htmltag84 }\htmlrtf {\htmlrtf0 " & Title & vbCrLf
'objfile.write "{\*\htmltag92 }\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 Office: " & PhoneNumber & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

If blnVoice = 0 Then
objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 Direct: " & strIPPhone & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf
End if

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 Fax   : " & FaxNumber & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f4 \htmlrtf0 " & StreetAddress & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
objfile.write vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana color=#000000 size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 "  & Town & ", " & State & " " & ZipCode & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf


objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana size=2>}\htmlrtf {\f0 \fs20 \htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag84 <A \par href=" & aQuote & web_address & aQuote & ">}\htmlrtf {\field{\*\fldinst{HYPERLINK " & aQuote & web_address & aQuote & "}}{\fldrslt\cf1\ul \htmlrtf0" & web_address & "\htmlrtf }\htmlrtf0 \htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag92 </A>}" & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag116 <BR>}\htmlrtf \line " & vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag0 \par }" & vbCrLf

objfile.write "{\*\htmltag96 <DIV align=left>}\htmlrtf {\htmlrtf0 {\*\htmltag64}\htmlrtf {\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag148 <FONT face=Verdana size=1>}\htmlrtf {\f0 \fs16 \htmlrtf0" & strDisclaimer & vbCrLf
objfile.write "{\*\htmltag156 </FONT>}\htmlrtf }\htmlrtf0 " & vbCrLf
objfile.write "{\*\htmltag84 &nbsp;}\htmlrtf \'a0\htmlrtf0 {\*\htmltag72}\htmlrtf\par}\htmlrtf0" & vbCrLf
'objfile.write vbCrLf
objfile.write "{\*\htmltag104 </DIV>}\htmlrtf }\htmlrtf0 " & vbCrLf

objfile.write "{\*\htmltag58 </BODY>}" & vbCrLf
objfile.write "{\*\htmltag27 </HTML>}" & vbCrLf
objfile.write "{\*\htmltag3 \par }}" & vbCrLf

objFile.Close
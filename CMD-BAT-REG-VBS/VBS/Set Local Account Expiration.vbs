'==========================================================================
'
' NAME: Set Local Account Expiration Date
'
' AUTHOR: Justin Baits
' DATE  : 12/2/2011
'
' Description: This script will prompt for a local user name and a desired
' date for that account to expire. It then sets that account to expire on
' the provided date.
'
'
' Comments: This script must be run from an elevated command prompt.
' If an invalid user name is entered a windows script error will tell you 
' "The user name could not be found"
'==========================================================================

Option Explicit

Dim objNetwork, strComputer, strName, objUser, dtmDate, strDate, Cancel


' Retrieve local computer name.
Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName


' Prompt for local user name.
Call GetUserName
Public Function GetUserName
   strName = InputBox("Enter user name to set expiration date", "User Name")
   If IsEmpty(strName) Then
       Cancel = MsgBox("Are you sure you want to exit ?", 68, "Exit Application?")
       If Cancel = vbYes Then
           WScript.Quit
       ElseIf Cancel = vbNo Then
           Call GetUserName
       End If
   ElseIf Len(strName) = 0 Then
       MsgBox "Enter a valid user name", 64, "Error"
       Call GetUserName
	End If
End Function	


' Prompt for account expiration date.
Call GetDate
Public Function GetDate
   strDate = InputBox("Enter desired expiration date MM/DD/YYYY", "Expiration Date")
   If IsEmpty(strDate) Then
       Cancel = MsgBox("Are you sure you want to exit ?", 68, "Exit Application?")
       If Cancel = vbYes Then
           WScript.Quit
       ElseIf Cancel = vbNo Then
           Call GetDate
       End If
   ElseIf Len(strDate) = 0 Then
       MsgBox "Enter a valid date", 64, "Error"
       Call GetDate
	End If
End Function
	
	
' Bind to local user object.
Set objUser = GetObject("WinNT://" & strComputer & "/" & strName & ",user")


' Convert date input to Date type
dtmDate = CDate(strDate)


' Assign expiration date.
objUser.AccountExpirationDate = dtmDate
objUser.SetInfo

'Tell user script is complete
MsgBox "Expiration date has been set", 64, "Success"
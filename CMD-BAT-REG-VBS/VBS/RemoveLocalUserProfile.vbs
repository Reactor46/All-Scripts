'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

On Error Resume Next

Dim intDay
Dim strComputer
Dim strFilter,objFSO,objWMIService
Dim colProfiles,objProfile,dtmLastUseTime,intUnusedDay
Dim result

intDay = InputBox("Please intput a number in the box,then the user profile not used for more than that number of days will be deleted.")
If IsEmpty(intDay) = True Then
	WScript.Quit
ElseIf CInt(intDay) < 0 Then
	WScript.Echo "Input incorrect, please follow the prompt and enter the number."	
Else
	strComputer = "."
	strFilter = "SID Like ""S-1-5-21%"" And Not LocalPath Like ""%Administrator%"""
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2") 
	Set colProfiles = objWMIService.ExecQuery("Select * From Win32_UserProfile Where " & strFilter)
	
	If Not colProfiles Is Nothing Then
		For Each objProfile in colProfiles
			'Converting a Standard Date to a WMI Date-Time Format
			dtmLastUseTime = CDate(Mid(objProfile.LastUseTime, 5, 2) & "/" & Mid(objProfile.LastUseTime, 7, 2) & "/" & Left(objProfile.LastUseTime, 4) & " " & Mid (objProfile.LastUseTime, 9, 2) & ":" & Mid(objProfile.LastUseTime, 11, 2) & ":" & Mid(objProfile.LastUseTime, 13, 2))
			intUnusedDay = DateDiff("d", dtmLastUseTime, Date)

			If intUnusedDay >= CInt(intDay) Then
				result = MsgBox("Are you sure you want to delete '" & objProfile.LocalPath & "' profile?", vbYesNo, "Delete UserProfile")
				If result = vbYes Then
					objProfile.Delete_
					WScript.Echo "Delete profile '" & objProfile.LocalPath & "' successfully."
				End If
			End If
		Next
	Else
		WScript.Echo "The item not found."
	End If	
End If
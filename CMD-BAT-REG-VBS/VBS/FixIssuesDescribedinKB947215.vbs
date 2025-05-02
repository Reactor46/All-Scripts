'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

Option Explicit

Dim ObjReg : Set ObjReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Dim ObjFSO : Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Dim ObjShell : Set ObjShell = CreateObject("WScript.Shell")

Function RepairUserAccountProfile(sID)
	On Error Resume Next
	 
    Dim HKEY_LOCAL_MACHINE : HKEY_LOCAL_MACHINE = &H80000002
    Dim KeyPath            : KeyPath = "Software\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Dim FullKeyPath        : FullKeyPath = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\"
    Dim ValueName          : ValueName = "ProfileImagePath"
	Dim subKeys , subKey , profileImagePath, strSIDs, arrSIDs, sidRegKeyCount, sidRegKey, sidRegKeyBak

	'	Check if a ¡°NT AUTHORITY¡± user SID is specified
	'	Omit bellow SIDs:
	'	s-1-5-18 for NT AUTHORITY\SYSTEM
	'	s-1-5-19 for NT AUTHORITY\LOCAL SERVICE
	'	s-1-5-20 for NT AUTHORITY\NETWORK SERVICE
    If sID = "S-1-5-18" Or sID = "S-1-5-19" Or sID = "S-1-5-20" Then
    	WScript.Echo "The SID: [" & sID & "] is a NT AUTHORITY account and cannot be repaired here!"
    	Exit Function 
    End If

    ObjReg.EnumKey HKEY_LOCAL_MACHINE, KeyPath, subKeys
    strSIDs = ""
    
    For Each subKey In subKeys
    	If InStr(subKey,sID) > 0 Then
    		strSIDs = strSIDs & "," & subKey
    	End If
    Next
    
    ' The inputted SID exists
    If strSIDs <> "" Then 
    	strSIDs = Mid(strSIDs,2)
    	arrSIDs = Split(strSIDs,",")
    Else 
    	WScript.Echo "The SID you input does not exist or is valid."
    	Exit Function
    End If 
    
    sidRegKeyCount = UBound(arrSIDs) - LBound(arrSIDs)+ 1
    
    ' There are too many SID backups for the inputted SID, cannot determine how to modify it.
    If sidRegKeyCount > 2 Then 
    	WScript.Echo "There are too many registry key backups for the SID [" & sID & "].It cannot be repaired."
    	Exit Function    
    End If 
    
    ' There is only one registry key starting with inputted SID and ends with .bak
    If sidRegKeyCount = 1 Then 
    	If arrSIDs(0) = sID Then 
    		WScript.Echo "There is no registry key backup for the user profile of SID [" & sID & "]."
    		Exit Function
    	Else 
    		ObjReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, KeyPath & "\" & arrSIDs(0), ValueName, profileImagePath
    		
    		If ObjFSO.FolderExists(profileImagePath) Then
				ObjShell.Run "cmd /c REG COPY """ & FullKeyPath & arrSIDs(0) &"""  """ & FullKeyPath & sID & """ /f",0,True
				ObjShell.Run "cmd /c REG DELETE """ & FullKeyPath & arrSIDs(0) &""" /f",0,True
    		Else 
    			WScript.Echo "The user profile for SID  [" & sID & "] cannot be repaired." 
    			Exit Function 
    		End If 
    	End If
    ' There are two registry keys starting with inputted SID and one of them ended with .bak 
    Else 
    	If arrSIDs(0) = sID Then 
    		sidRegKey = arrSIDs(0)
    		sidRegKeyBak = arrSIDs(1)
    	ElseIf arrSIDs(0) = sID Then 
    		sidRegKey = arrSIDs(0)
    		sidRegKeyBak = arrSIDs(1)
    	Else 
    		WScript.Echo "There are too many registry key backups for the user profile of SID [" & sID & "]."
    		Exit Function 
    	End If 
    	
    	ObjReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, KeyPath & "\" & sidRegKeyBak, ValueName, profileImagePath
    		
		If ObjFSO.FolderExists(profileImagePath) Then
			ObjShell.Run "cmd /c REG COPY """ & FullKeyPath & sidRegKey &"""  """ & FullKeyPath & sidRegKey & "_Temp"" /f" ,0,True 
			ObjShell.Run "cmd /c REG DELETE """ & FullKeyPath & sidRegKey &""" /f" ,0,True 
			ObjShell.Run "cmd /c REG COPY """ & FullKeyPath & sidRegKeyBak &"""  """ & FullKeyPath & sidRegKeyBak & "_Temp"" /f",0,True
			ObjShell.Run "cmd /c REG DELETE """ & FullKeyPath & sidRegKeyBak &""" /f",0,True
			ObjShell.Run "cmd /c REG COPY """ & FullKeyPath & sidRegKey & "_Temp"" """ & FullKeyPath & sidRegKey & ".bak"" /f",0,True
			ObjShell.Run "cmd /c REG DELETE """ & FullKeyPath & sidRegKey & "_Temp"" /f",0,True
			ObjShell.Run "cmd /c REG COPY """ & FullKeyPath & sidRegKeyBak & "_Temp"" """ & FullKeyPath & sidRegKey & """ /f",0,True
			ObjShell.Run "cmd /c REG DELETE """ & FullKeyPath & sidRegKeyBak & "_Temp""  /f",0,True
		Else 
			WScript.Echo "The user profile for SID  [" & sID & "] cannot be repaired."
			Exit Function 
		End If 
    End If 
    
	ObjReg.SetDWORDValue   HKEY_LOCAL_MACHINE,KeyPath & "\" & sID, "RefCount",0
	ObjReg.SetDWORDValue   HKEY_LOCAL_MACHINE,KeyPath & "\" & sID, "State",0
	
	WScript.Echo "The user account profile for SID [" & sID & "] was repaired successfully." 
	WScript.Echo "You need restart your computer to have this change to take effect."
End Function

Function DeleteInvalidSIDRegKey()
	On Error Resume Next
	 
    Dim HKEY_LOCAL_MACHINE : HKEY_LOCAL_MACHINE = &H80000002
    Dim KeyPath            : KeyPath = "Software\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Dim ValueName          : ValueName = "ProfileImagePath"
	Dim subKeys , subKey , profileImagePath
	
    ObjReg.EnumKey HKEY_LOCAL_MACHINE, KeyPath, subKeys

    For Each subKey In subKeys
	    '	Check if a ¡°NT AUTHORITY¡± user SID is specified
		'	Omit bellow SIDs:
		'	s-1-5-18 for NT AUTHORITY\SYSTEM
		'	s-1-5-19 for NT AUTHORITY\LOCAL SERVICE
		'	s-1-5-20 for NT AUTHORITY\NETWORK SERVICE
        If subKey <> "S-1-5-18" And subKey <> "S-1-5-19" And subKey <> "S-1-5-20" Then 
	        ObjReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, KeyPath & "\" & subKey, ValueName, profileImagePath
	        
	        ' If a SID with a invalid profileImagePath, this SID is invalid, and it need to be deleted
	        If Not ObjFSO.FolderExists(profileImagePath) Then
	            ObjReg.DeleteKey HKEY_LOCAL_MACHINE, KeyPath & "\" & subKey
	            WScript.Echo "Invalid SID: [" & subKey & "] was deleted successfully!"
	        End If
        End If 
    Next
End Function


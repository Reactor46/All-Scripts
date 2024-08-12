'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 08/01/2009
' Description : This Script will Export all the Last logon sessions 
'               of all the users that connected to a remote computer
'=*=*=*=*=*=*=*=*=*=*=*=
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = WScript.Arguments(0)
Set pbjReg  = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv" )
strKeyPath  = "SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList"
pbjReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
For Each subkey In arrSubKeys
	ret = pbjReg.GetStringvalue(HKEY_LOCAL_MACHINE, strKeyPath & "\ProfileImagePath" , sValue)
	If ret<>0 Then
	   Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv" )
	   strKeyPath = "SOFTWARE\MICROSOFT\Windows NT\CurrentVersion\ProfileList\" & subkey
	   strValueName = "ProfileImagePath"
	   Return = objReg.GetExpandedStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
	   strValueName = "State"
	   Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngState)
	   If (lngState > 0) Then
	   	  UserName = Split(strValue,"\")(2)
	      strValueName = "ProfileLoadTimeHigh" ' Calculate Load Time High
	      Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngHighValue)
	      strValueName = "ProfileLoadTimeLow" ' Calculate Load Time Low
	      Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngLowValue)
	      strValueName = "OptimizedLogonStatus" ' Check if User Logged on Localy or by Network
	      Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngLogonStatus)
	      strValueName = "RefCount" ' Check if User is currently logged on
	      Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngCurLogon)
	      strValueName = "Flags" ' Check if User is a script
	      Return = objReg.GetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,lngFlag)
	      If (TypeName(lngFlag) <> "Null") And (lngFlag = 1) Then
	      		UserName = UserName & ",Script," ' User Logged on by Service
	      ElseIf TypeName(lngLogonStatus) <> "Null" Then
	      		UserName = UserName & ",Locally," ' User Logged on Locally
	      Else
	      	UserName = UserName & ",Network," ' User Logged on by Network (runas, UNC, Script, Ext.)
	      End If
	      
	      If lngCurLogon = 1 Then
	      	 UserName = UserName & "Logged on," ' User is currently logged on
	      Else
	      	 UserName = UserName & "Logged off," ' User is currently logged off
	      End If
	      
	  	  If typename(lngHighValue) <> "Null" Then
	        NanoSecs = (lngHighValue * 2 ^ 32 + lngLowValue)
	        '' /* Returns UTC (Universal Coordinated Time, aka GMT) time */
	        DT = #1/1/1601# + (NanoSecs / 600000000 / 1440)
	        WScript.Echo UserName & CDate(DT)
		  End If
	    End If
	End if
Next
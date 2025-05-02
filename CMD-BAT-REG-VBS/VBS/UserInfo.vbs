'-----------------------------------------------------------------------------
' Define Constants
'-----------------------------------------------------------------------------
Const ADS_UF_DONT_EXPIRE_PASSWORD = &h10000
Const ADS_UF_SMARTCARD_REQUIRED = &h40000 
Const SEC_IN_MIN = 60
Const SEC_IN_DAY = 86400
Const MIN_IN_DAY = 1440
Const ADS_SCOPE_SUBTREE = 60
Const IAAP_EXPIRATION_WARNING_DAYS = 70

'-----------------------------------------------------------------------------
' Define Variables
'-----------------------------------------------------------------------------
Dim objShell, objNet, objRootDSE, objInfoDict, objEnv, objADInfo
Dim strADsConfPath, strRootDSE, strDN

'Dim strLogFile, strUser, strUserName, aryADUserInfo
'Dim returnCode, objLog, strComputer, strMOTDSrc, strMOTDLink

set objShell = CreateObject("WScript.Shell")
set objNet = CreateObject("WScript.Network")
set objRootDSE = GetObject("LDAP://rootDSE")
Set objInfoDict = CreateObject("Scripting.Dictionary")
Set objEnv = objShell.Environment("Process")
Set objADInfo = CreateObject("ADSystemInfo")


' Bind to the configuration to get Domain Controlers
strADsConfPath = "LDAP://" & objRootDSE.Get("configurationNamingContext")
' Bind to the default context for portability
strRootDSE = objRootDSE.Get("defaultNamingContext")

On Error Resume Next

objInfoDict.Add "Domain", UCase(objNet.UserDomain)
objInfoDict.Add "UserName", UCase(objNet.UserName)
objInfoDict.Add "OS", objEnv("OS")
objInfoDict.Add "EnvUserName", objEnv("USERNAME")

strDN = GetUserDN()

GetUserAccountInfo(strDN)

DisplayUserInfo()

WScript.Quit

'============  Functions and Subroutines ================
Private Sub DisplayUserInfo()
   Dim strMsg, strDebug, strKey
   Dim colKeys
   strMsg = "Please verify the following information.  If anything is " & _
            "incorrect, contact the help desk." & vbCrLf & vbCrLf & _
            "Display Name: " & vbTab & vbTab & objInfoDict.Item("displayName") & vbCrLf & _
            "Description: " & vbTab & vbTab & objInfoDict.Item("description") & vbCrLf & _
            "Department: " & vbTab & vbTab & objInfoDict.Item("physicalDeliveryOfficeName") & vbCrLf & _
            "Telephone Number: " & vbTab & objInfoDict.Item("telephoneNumber") & vbCrLf & vbCrLf & _
            "Failed Logon Attempts: " & vbTab & objInfoDict.Item("badPwdCount") & vbCrLf & _
            "Last Failed Logon: " & vbTab & vbTab & objInfoDict.Item("lastFailedLogin") & vbCrLf & _
            "Last Successful Logon: " & vbTab & objInfoDict.Item("lastLogin") & vbCrLf & _
            "Last Password Change: " & vbTab & objInfoDict.Item("passwordLastChanged") & vbCrLf & _
            "Password Expires: " & vbTab & vbTab & objInfoDict.Item("pwdExpires") & vbCrLf

   If objInfoDict.Item("IAAP Warning") <> "" Then
      strMsg = strMsg & vbCrLf & vbCrLf & objInfoDict.Item("IAAP Warning")
   End If
   
   If objEnv("DEBUG") <> "" Then
      colKeys = objInfoDict.Keys
      For Each strKey in colKeys
         strDebug = strDebug & strKey & ": " & objInfoDict(strKey) & vbCrLf
      Next
      objShell.Popup strDebug, 0, "*** Debugging Information ***"
   End If
   
   objShell.Popup strMsg, 60, "User Information for: " & objInfoDict.Item("displayName")
End Sub  ' DisplayUserInfo

Private Function GetUserDN()
   Dim objConnection, objCommand, objRecordSet, objUser
   
   On Error Resume Next

   Const ADS_SCOPE_SUBTREE = 2

   Set objConnection = CreateObject("ADODB.Connection")
   Set objCommand =   CreateObject("ADODB.Command")
   objConnection.Provider = "ADsDSOObject"
   objConnection.Open "Active Directory Provider"
   Set objCommand.ActiveConnection = objConnection
   objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
   objCommand.CommandText = _
       "SELECT ADsPath FROM 'LDAP://" & strRootDSE & _
       "' WHERE objectCategory='user' AND samAccountName='" & objInfoDict.Item("EnvUserName") & "'"  
   Set objRecordSet = objCommand.Execute

   strPath = objRecordSet.Fields("ADsPath").Value
   Set objUser = GetObject(strPath)

   GetUserDN = objUser.distinguishedName
   
   Set objRecordSet = Nothing
   objConnection.Close
   Set objConnection = Nothing
   Set objUser = Nothing
End Function  ' GetUserDN

' This subroutine gathers more information that is displayed.  This is done to
' consolidate and document the information that can be displayed in the future.
Private Sub GetUserAccountInfo(strDN)
   Dim objDomainNT, objUser, objADs, objDC, objWMIService, objTimeZones, objADUser
   Dim dtTemp
   Dim strLogonServer
   Dim iTimeZoneBias, iDaylightBias
   
   On Error Resume Next
   
   If InStr(1, strDN, "/") Then
      strDN = Replace(strDN, "/", "\/")
   End IF
   strLogonServer = Replace(objEnv("LOGONSERVER"), "\", "")
   
   Set objDomainNT = GetObject("WinNT://" & objInfoDict.Item("Domain"))
   Set objUser = GetObject("LDAP://" & strDN)
   Set objADs = GetObject("LDAP://" & strRootDSE)
   Set objDC = GetObject("LDAP://" & strLogonServer & "/" & strDN)
   'Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
   'Set objTimeZones = objWMIService.ExecQuery("Select * From Win32_TimeZone")
   
   ' Due to the number of domain controllers, only query the DC currently
   ' connected to for last information
   '-------------------------------------------------------------------------
   objInfoDict.Add "lastFailedLogin", objDC.lastFailedLogin
   objInfoDict.Add "lastLogin", objDC.lastLogin
   objInfoDict.Add "badPwdCount", objDC.badPwdCount

   ' Get Timezone Information...
   '-------------------------------------------------------------------------
   strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colTimeZones = objWMIService.ExecQuery("Select * From Win32_TimeZone")

For Each objTimeZone in objTimeZones
    intTimeZoneBias = objTimeZone.Bias
    intDaylightBias = objTimeZone.DaylightBias
   Next
   objInfoDict.Add "TZ Bias", intTimeZoneBias
   objInfoDict.Add "TZ Daylight Bias", intDaylightBias

   ' Get NT Domain Information...
   '-------------------------------------------------------------------------
   With objDomainNT
      objInfoDict.Add "maxPasswordAge", .Get("maxPasswordAge")  ' In seconds
      objInfoDict.Add "maxPasswordAgeDays", (.Get("maxPasswordAge") / SEC_IN_DAY)  ' Convert seconds to days
      objInfoDict.Add "minPasswordAgeSeconds", .Get("minPasswordAge")  ' In seconds
      objInfoDict.Add "minPasswordAgeDays", (.Get("minPasswordAge") / SEC_IN_DAY)  ' Convert seconds to days
      objInfoDict.Add "lockoutObservationInterval", .Get("lockoutObservationInterval")  ' In seconds
      objInfoDict.Add "lockoutObservationIntervalDays", (.Get("lockoutObservationInterval") / SEC_IN_DAY)  ' Convert seconds to days
      objInfoDict.Add "autoUnlockInterval", .Get("autoUnlockInterval")  ' In seconds
      If objInfoDict.Item("autoUnlockUnterval") <> -1 Then
         objInfoDict.Add "autoUnlockIntervalMin", (.Get("autoUnlockInterval") / SEC_IN_MIN)  ' Convert seconds to minutes
      Else
         objInfoDict.Add "autoUnlockIntervalMin", "Manually Unlocked by Administrator"
      End IF   
   End With

   ' Get Main User Information...
   '-------------------------------------------------------------------------
   With objUser
      objInfoDict.Add "displayName", .Get("displayName")
      objInfoDict.Add "givenName", .Get("givenName")
      objInfoDict.Add "initials", .Get("initials")
      objInfoDict.Add "sn", .Get("sn")
      objInfoDict.Add "telephoneNumber", .Get("telephoneNumber")
      objInfoDict.Add "sAMAccountName", .Get("sAMAccountName") ' Logon Name
      objInfoDict.Add "description", .Get("description")
      objInfoDict.Add "physicalDeliveryOfficeName", .Get("physicalDeliveryOfficeName")  ' Department
      dtTemp = .Get("whenCreated")
      dtTemp = DateAdd("n", intTimeZoneBias, dtTemp)
      objInfoDict.Add "whenCreated", DateValue(dtTemp) & " at " & TimeValue(dtTemp)
      dtTemp1 = .Get("PasswordLastChanged")  '******** whenChanged
      dtTemp = DateAdd("n", intTimeZoneBias, dtTemp)
      objInfoDict.Add "whenchanged", DateValue(dtTemp) & " at " & TimeValue(dtTemp)  '******* pwdLastSet to whenchanged
      objInfoDict.Add "userAccountControl", .Get("userAccountControl")
      dtTemp = .PasswordLastChanged
      dtTemp = DateAdd("n", intTimeZoneBias, dtTemp)

      If dtTemp = "" Then
         objInfoDict.Add "passwordAgeDays", ""
         objInfoDict.Add "passwordLastChanged", "Not Available"
         objInfoDict.Add "passwordExpired", "Not Available"
      Else
         objInfoDict.Add "passwordAgeDays", int(now - dtTemp)
         If objInfoDict.Item("maxPasswordAgeDays") >= objInfoDict.Item("passwordAgeDays") Then
            objInfoDict.Add "passwordExpired", "No"
         Else
            objInfoDict.Add "passwordExpired", "Yes"
         End If
         objInfoDict.Add "passwordLastChanged", DateValue(dtTemp) & " at " & TimeValue(dtTemp)
      End If
      If objInfoDict.Item("userAccountControl") And ADS_UF_DONT_EXPIRE_PASSWORD Then
         objInfoDict.Add "pwdExpires", "Password does not expire"
         objInfoDict.Add "pwdNeverExpires", "Yes"
  end if
     if objInfoDict.Item("userAccountControl") And ADS_UF_SMARTCARD_REQUIRED Then
         objInfoDict.Add "pwdExpires", "Account Smartcard Required"

     else
         objInfoDict.Add "pwdExpires", DateValue(dtTemp + 90) & _
                                       " at " & TimeValue(dtTemp) & " (" & _
                                       int((dtTemp + 90) - now) & " days from today" & ")." ' dtTemp = passwordLastChanged
         objInfoDict.Add "pwdNeverExpires", "No"
     End If
      objInfoDict.Add "mailNickname", .Get("mailNickname")
      If objInfoDict.Item("mailNickname") = "" Then
         objInfoDict.Add "exchange", "No"
      Else
         objInfoDict.Add "exchange", "Yes"
      End If
   End With

   ' Get AD User Information...
   '-------------------------------------------------------------------------
   With objADs
      objInfoDict.Add "minPwdLength", .Get("minPwdLength")
      objInfoDict.Add "pwdHistoryLength", .Get("pwdHistoryLength")
      objInfoDict.Add "pwdProperties", .Get("pwdProperties")
      objInfoDict.Add "lockoutThreshold", .Get("lockoutThreshold")
   End With
   
   ' Get IAAP Information...
   '-------------------------------------------------------------------------
   Set objADUser = GetObject("LDAP://" & objADInfo.UserName)
   objInfoDict.Add "accountExpirationDate", (objADUser.AccountExpirationDate - 1)
   dtToday = Date()  ' Get the current date
   dtAcctExp = objInfoDict.Item("accountExpirationDate")
   dtDiff = datediff("d", dtToday, dtAcctExp)  ' Comparing the account expiration date and the current date

   'Calculations used to determine if the account will expire soon
   If dtDiff >= 0 and dtDiff <= IAAP_EXPIRATION_WARNING_DAYS Then ' Set it to # days to set off the alarm.
      objInfoDict.Add "IAAP Warning", _
                      "Warning, your account will expire in " & dtDiff & _
                      " days (" & dtAcctExp & ")." & vbCrLf & _
                      "Please take the current Information Assurance Awareness CBT course located" & vbCrLf & _
	                   "on the Air Force Portal (IT E-Learning link).  If your training is up to date" & vbCrLf & _
                      "(less than 1 year old), please contact help desk for further guidance."
   Else
      objInfoDict.Add "IAAP Warning", ""
   End If
   
End Sub  ' GetUserAccountInfo

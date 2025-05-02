' *********************************************
' Labguage VBS
' Create a Computer Account For a Specific User
' By Roma 06/01/2012
' v1.0.0.0
' *********************************************

' Some var's
Dim strComputer,strComputerUser,objRootDSE,objContainer,objComputer
Dim objSecurityDescriptor,objDACL,objACE1,objACE2,objACE3,objACE4
Dim objACE5,objACE6,objACE7,objACE8,objACE9

' computer name InputBox
strComputer = InputBox("Example: Computer-Account-Name","Create a Computer Account")
	If strComputer = "" Then
		WScript.Quit
End If

' user name InputBox
strComputerUser = InputBox("Example: domainname\username","Specific User")
	If strComputerUser = "" Then
		WScript.Quit
End If

' some constance
Const ADS_UF_PASSWD_NOTREQD = &h0020
Const ADS_UF_WORKSTATION_TRUST_ACCOUNT = &h1000
Const ADS_ACETYPE_ACCESS_ALLOWED = &h0
Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &h5
Const ADS_FLAG_OBJECT_TYPE_PRESENT = &h1
Const ADS_RIGHT_GENERIC_READ = &h80000000
Const ADS_RIGHT_DS_SELF = &h8
Const ADS_RIGHT_DS_WRITE_PROP = &h20
Const ADS_RIGHT_DS_CONTROL_ACCESS = &h100

Const ALLOWED_TO_AUTHENTICATE = _
    "{68B1D179-0D15-4d4f-AB71-46152E79A7BC}"
Const RECEIVE_AS = "{AB721A56-1E2f-11D0-9819-00AA0040529B}"
Const SEND_AS = "{AB721A54-1E2f-11D0-9819-00AA0040529B}"
Const USER_CHANGE_PASSWORD = _
    "{AB721A53-1E2f-11D0-9819-00AA0040529b}"
Const USER_FORCE_CHANGE_PASSWORD = _
    "{00299570-246D-11D0-A768-00AA006E0529}"
Const USER_ACCOUNT_RESTRICTIONS = _
    "{4C164200-20C0-11D0-A768-00AA006E0529}"
Const VALIDATED_DNS_HOST_NAME = _
    "{72E39547-7B18-11D1-ADEF-00C04FD8D5CD}"
Const VALIDATED_SPN = "{F3A64788-5306-11D1-A9C5-0000F80367C1}"

' get objects
Set objRootDSE = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://cn=Computers," & _
    objRootDSE.Get("defaultNamingContext"))
 
Set objComputer = objContainer.Create _
    ("Computer", "cn=" & strComputer)
objComputer.Put "sAMAccountName", strComputer & "$"
objComputer.Put "userAccountControl", _
    ADS_UF_PASSWD_NOTREQD Or ADS_UF_WORKSTATION_TRUST_ACCOUNT
objComputer.SetInfo
 
Set objSecurityDescriptor = objComputer.Get("ntSecurityDescriptor")
Set objDACL = objSecurityDescriptor.DiscretionaryAcl
 
Set objACE1 = CreateObject("AccessControlEntry")
objACE1.Trustee    = strComputerUser
objACE1.AccessMask = ADS_RIGHT_GENERIC_READ
objACE1.AceFlags   = 0
objACE1.AceType    = ADS_ACETYPE_ACCESS_ALLOWED
 
Set objACE2 = CreateObject("AccessControlEntry")
objACE2.Trustee    = strComputerUser
objACE2.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
objACE2.AceFlags   = 0
objACE2.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE2.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE2.ObjectType = ALLOWED_TO_AUTHENTICATE
 
Set objACE3 = CreateObject("AccessControlEntry")
objACE3.Trustee    = strComputerUser
objACE3.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
objACE3.AceFlags   = 0
objACE3.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE3.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE3.ObjectType = RECEIVE_AS
 
Set objACE4 = CreateObject("AccessControlEntry")
objACE4.Trustee    = strComputerUser
objACE4.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
objACE4.AceFlags   = 0
objACE4.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE4.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE4.ObjectType = SEND_AS
 
Set objACE5 = CreateObject("AccessControlEntry")
objACE5.Trustee    = strComputerUser
objACE5.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
objACE5.AceFlags   = 0
objACE5.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE5.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE5.ObjectType = USER_CHANGE_PASSWORD
 
Set objACE6 = CreateObject("AccessControlEntry")
objACE6.Trustee    = strComputerUser
objACE6.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
objACE6.AceFlags   = 0
objACE6.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE6.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE6.ObjectType = USER_FORCE_CHANGE_PASSWORD
 
Set objACE7 = CreateObject("AccessControlEntry")
objACE7.Trustee    = strComputerUser
objACE7.AccessMask = ADS_RIGHT_DS_WRITE_PROP
objACE7.AceFlags   = 0
objACE7.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE7.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE7.ObjectType = USER_ACCOUNT_RESTRICTIONS
 
Set objACE8 = CreateObject("AccessControlEntry")
objACE8.Trustee    = strComputerUser
objACE8.AccessMask = ADS_RIGHT_DS_SELF
objACE8.AceFlags   = 0
objACE8.AceType    = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE8.Flags      = ADS_FLAG_OBJECT_TYPE_PRESENT
objACE8.ObjectType = VALIDATED_DNS_HOST_NAME
 
Set objACE9 = CreateObject("AccessControlEntry")
objACE9.Trustee    = strComputerUser
objACE9.AccessMask = ADS_RIGHT_DS_SELF
objACE9.AceFlags   = 0
objACE9.AceType  = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
objACE9.Flags  =  ADS_FLAG_OBJECT_TYPE_PRESENT
objACE9.ObjectType = VALIDATED_SPN
 
objDACL.AddAce objACE1
objDACL.AddAce objACE2
objDACL.AddAce objACE3
objDACL.AddAce objACE4
objDACL.AddAce objACE5
objDACL.AddAce objACE6
objDACL.AddAce objACE7
objDACL.AddAce objACE8
objDACL.AddAce objACE9
 
objSecurityDescriptor.DiscretionaryAcl = objDACL
objComputer.Put "ntSecurityDescriptor", objSecurityDescriptor
objComputer.SetInfo

' memory cleaning
Set strComputer = Nothing
Set strComputerUser = Nothing
Set objRootDSE = Nothing
Set objContainer = Nothing
Set objComputer = Nothing
Set objSecurityDescriptor = Nothing
Set objDACL = Nothing
Set objACE1 = Nothing
Set objACE2 = Nothing
Set objACE3 = Nothing
Set objACE4 = Nothing
Set objACE5 = Nothing
Set objACE6 = Nothing
Set objACE7 = Nothing
Set objACE8 = Nothing
Set objACE9 = Nothing
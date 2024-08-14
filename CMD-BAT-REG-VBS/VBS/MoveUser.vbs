'=========================================================================== 
'  MoveUser.vbs
'  Rob Greene of Technet created the original script
'  Ron Williams expanded it to have similar functionality to the previous Moveuser.exe 
'
'  This script has one hard coded variable in it that must be filled out. 
'  Line #28 - strDomainDN = The Target Account Domain's Distinguished Name.  
'=========================================================================== 
Option Explicit 
DIM strComputer, strSourceAcct, strSourceAcctDomain, strTargetAcct 
DIM strTargetAcctDomain, strTargetAcctSID, objComputer, objOs 
DIM objProfile, objCommand, objRecordSet, objConnection, objWMIService, objSID 
DIM dtStart, colProfiles, oSID, oUsr, strDomainDN, strWmios, strOsQuery
DIM Revision, IssueAuthorities(11), strSDDL, subAuthorities
DIM strDomAcctLength, strCompName, strSlashPos, strContinue
DIM strDomainAcct, strLocalAcct, strLocalAcctLength, strArg3
DIM strKeepuser, strArg4, colOperatingSystems, strSrcDom, strOSVer

CONST ADS_SCOPE_SUBTREE=2 

'=========================================================================== 
'  This script has one hard coded variable in it that must be filled out. 
'  strDomainDN = The Target Account Domain's Distinguished Name.  
'  This is done for the LDAP query to be built to find the target account's SID
'  Sample 1:  strDomainDN="dc=microsoft,dc=com"   
'  Sample 2:  strDomainDN="dc=dept,dc=university,dc=edu"   

strDomainDN="dc=mydomain,dc=lk"
'=========================================================================== 

If WScript.Arguments.Count >= 2 Then
strLocalAcct = WScript.Arguments.Item(0)
strDomainAcct = WScript.Arguments.Item(1)
 Else
  Call Syntax
End If

'Process arg 1
If InStr (strLocalAcct, "\") = 0 Then
 Set strCompName = WScript.CreateObject("WScript.Network")
 strSourceAcctDomain = strCompName.ComputerName
 strSourceAcct = strLocalAcct
 Else
   strLocalAcctLength = Len(strLocalAcct)
   strSlashPos = InStr (strLocalAcct, "\")
   strSourceAcctDomain = Left(strLocalAcct,strSlashPos - 1)
   strSourceAcct = Right(strLocalAcct,strLocalAcctLength - strSlashPos)
   strSrcDom = "Yes"
End If

'Process arg 2
If InStr (strDomainAcct, "\") = 0 Then
 Set strCompName = WScript.CreateObject("WScript.Network")
 strTargetAcctDomain = strCompName.ComputerName
 strTargetAcct = strDomainAcct
 Else
	 strDomAcctLength = Len(strDomainAcct)
	 strSlashPos = InStr (strDomainAcct, "\")
	 strTargetAcctDomain = Left(strDomainAcct,strSlashPos - 1)
	 strTargetAcct = Right(strDomainAcct,strDomAcctLength - strSlashPos)
End If

'Process arg 3
If WScript.Arguments.Count >= 3 Then
strArg3 = WScript.Arguments.Item(2)
  If Left(strArg3,2)="/c" Then
    strComputer = Right(strArg3, Len(strArg3)-3)      
    strKeepuser="no"
    Elseif strArg3="/k" Then
      strComputer ="."
      strKeepuser="yes"
    Else
      strComputer ="."
      strKeepuser="no"
    End If
Else
strComputer ="."
strKeepuser="no"
End If

'Process arg 4
If WScript.Arguments.Count = 4 Then
strArg4 = WScript.Arguments.Item(3)
  If Left(strArg4,2)="/c" Then
    strComputer = Right(strArg4, Len(strArg4)-3)      
    Elseif strArg4="/k" Then
      strKeepuser="yes"
    Else
      strKeepuser="no"
    End If
End If

If WScript.Arguments.Count > 4 Then
  Call Syntax
End If
If LCase(strLocalAcct)="administrator" Then
      strKeepuser="yes"
End If
If strSrcDom = "Yes" Then
      strKeepuser="yes"
End If
If strDomainDN = "dc=contoso,dc=com" Then
    Call Change_DDN
End If
'==========OS/SP Check============='
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 
strOsQuery = "Select * from Win32_OperatingSystem" 
Set colOperatingSystems = objWMIService.ExecQuery(strOsQuery) 
   For Each objOs in colOperatingSystems 
        strWmios = objOs.Caption & " - " & objOs.Version
        strOSVer = objOs.Version
        If Left(strOSVer,1) < 6 Then   
        'If InStr (strWmios, "XP") Then
            WScript.Echo "This script only works with Windows Vista SP1 and above."
            WScript.Quit
          Else If InStr (strWmios, "Vista") Then
              If strOSVer = "6.0.6000" Then
              WScript.Echo "Service Pack 1 Not Installed."
              WScript.Echo "Service Pack 1 or higher is required" & VBNewLine _
                         & "for this script to function."
              WScript.Quit
              End If
'          Else
'            WScript.Echo "This script only works with Windows Vista SP1 and above."
'            WScript.Quit
          End If
        End If
    Next 
'==========End OS/SP Check============='
strTargetAcctSID="" 
dtStart = TimeValue(Now()) 
Set objConnection = CreateObject("ADODB.Connection") 
objConnection.Open "Provider=ADsDSOObject;" 
Set objCommand = CreateObject("ADODB.Command") 
objCommand.ActiveConnection = objConnection 

' We need the proper Active Directory domain name where the user exists in a DN format.  You can 
' modify the strDomainDN variable to your Active Directory domain name in DN format. 

objCommand.CommandText = _ 
    "SELECT AdsPath, cn FROM 'LDAP:// " + strDomainDN + "' WHERE objectCategory = 'user'" & "And sAMAccountName= '" + strTargetAcct + "'" 
objcommand.Properties("searchscope") = ADS_SCOPE_SUBTREE  
Set objRecordSet = objCommand.Execute 
If objRecordset.RecordCount = 0 Then 
    WScript.Echo "sAMAccountName: " & strTargetAcct & " does not exist." 
ElseIf objRecordset.RecordCount > 1 Then 
    WScript.Echo "There is more than one account with the same sAMAccountName" 
Else 
    objRecordSet.MoveFirst 
    Do Until objRecordSet.EOF 
        Set Ousr = GetObject(objRecordSet.Fields("AdsPath").Value) 
        strTargetAcctSID = SDDL_SID(oUsr.Get("objectSID")) 
        objRecordSet.MoveNext 
    Loop 
objConnection.Close 

'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
Set colProfiles = objWMIService.ExecQuery("Select * from Win32_UserProfile") 
For Each objProfile in colProfiles 
    Set objSID = objWMIService.Get("Win32_SID.SID='" & objProfile.SID &"'") 


'    Testing to verify that the current profile handle is for 
'    the Source Account that we want to move to the domain user. 
if UCase(objsid.referencedDomainName + "\" + objsid.AccountName)= UCase(strSourceAcctDomain + "\" + strSourceAcct) Then 
    ' Making sure that the source profile is currently not in use.  If it is we will bail out. 
        If objProfile.RefCount < 1 Or IsNull(objProfile.RefCount) Then
            strContinue = MsgBox("Change Profile for:  " + strSourceAcctDomain + "\" + _
                          strSourceAcct + " to: " + strTargetAcctDomain + "\" + strTargetAcct + "?", _
                          vbYesNo + vbQuestion, "Move This Profile?") 
            If strContinue = 6 Then   '6=Yes, Move the Profile. 
              ' ChangeOwner method requires the String SID of Target Account and a Flag setting 
              ' Flag 1 = Change ownership of the source profile to target account even if the target account  already has a profile on the system. 
              ' Flag 2 = Delete the target account Profile and change ownership of the source user account profile to the target account. 
              ' To use the ChangeOwner method, both the source and target account profiles (If it exists) must not be loaded. 

                        ObjProfile.ChangeOwner strTargetAcctSID,1 
                        If strKeepUser="no" Then
                           If strComputer ="." Then
                              Set objComputer = GetObject("WinNT://" & strComputer & "")
                                  objComputer.Delete "user", strSourceAcct
                           End If
                        End If
                        WScript.Echo "Process Complete."
              ElseIf strContinue = 7 Then  '7=No, Cancel the move. 
                Wscript.Quit
              Else 
                Wscript.Quit

            End If 
     Else 
            Wscript.echo "Could not move the users profile, because " + _
            strSourceAcctDomain + "\" + strSourceAcct + " profile is currently loaded" 
        End If    
     End If 
Next 
End If 


Sub Init_IssueAuthorities( ) 
    'DIM IssueAuthorities(11) 
    IssueAuthorities(0) = "-0-0" 
    IssueAuthorities(1) = "-1-0" 
    IssueAuthorities(2) = "-2-0" 
    IssueAuthorities(3) = "-3-0" 
    IssueAuthorities(4) = "-4" 
    IssueAuthorities(5) = "-5" 
    IssueAuthorities(6) = "-?" 
    IssueAuthorities(7) = "-?" 
    IssueAuthorities(8) = "-?" 
    IssueAuthorities(9) = "-?" 
end sub 

Sub Syntax
  Wscript.Echo "Usage: Moveuser.vbs FromUser ToUser [/c:TargetComputer] [/k]" _
  & VBNewLine _
  & VBNewLine & "/c:TargetComputer will run against another system." _
  & VBNewLine & "/k indicates that if the Source user is a local account, do not delete the account after migration." _
  & VBNewLine _
  & VBNewLine & "Sample 1: Moveuser.vbs Fred Domain\Smithf" _
  & VBNewLine & "Sample 2: Moveuser.vbs Fred Domain\Smithf /k"
  Wscript.Quit
End Sub

Sub Change_DDN
  strContinue = MsgBox("This script has one hard coded variable in it that must be filled out." + (Chr(13)) _
              + "strDomainDN = The Target Account Domain's Distinguished Name." + (Chr(13)) _
              + "This is done to build the LDAP query that finds the target account's SID" + (Chr(13)) _
              + (Chr(13)) _
              + "You will find this variable on line# 28 set for the default 'sample'" + (Chr(13)) _
              + "domain name of contoso.com. Change it for your domain." + (Chr(13)) _
              + (Chr(13)) _
              + "Sample 1:  strDomainDN="+(Chr(34))+"dc=microsoft,dc=com"+(Chr(34)) + (Chr(13)) _
              + "Sample 2:  strDomainDN="+(Chr(34))+"dc=dept,dc=university,dc=edu"+(Chr(34)) _
              , vbOKOnly + vbInformation, "Variable Needs Edited.") 
  Wscript.Quit
End Sub

function SDDL_SID ( oSID ) 
    DIM Revision, SubAuthorities, strSDDL, IssueIndex, index, i, k, p2, subtotal 
    DIM j, dblSubAuth 
' 
' First byte is the revision value 
' 
    Revision = "1-5"
' 
' Second byte is the number of sub authorities in the 
' SID 
' 
    SubAuthorities = CInt(ascb(midb(oSID,2,1))) 
    strSDDL = "S-" & Revision 
    IssueIndex = CInt(ascb(midb(oSID,8,1))) 
' 
' BYtes 2 - 8 are the issuing authority structure 
' Currently these values are in the form: 
' { 0, 0, 0, 0, 0, X} 
' 
' We use this fact to retrieve byte number 8 as the index 
' then look up the authorities for an array of values 
' 
    strSDDL = strSDDL & IssueAuthorities(IssueIndex) 
' 
' The sub authorities start at byte number 9. The are 4 bytes long and 
' the number of them is stored in the Sub Authorities variable. 
' 
    index = 9 
    i = index 
    for k = 1 to SubAuthorities 
        ' 
        ' Very simple formula, the sub authorities are stored in the 
        ' following order: 
        ' Byte Index Starting Bit 
        ' Byte 0 - Index 0 
        ' Byte 1 - Index + 1 7 
        ' Byte 2 - Index + 2 15 
        ' Byte 3 - Index + 3 23 
        ' Bytes0 - 4 make a DWORD value in whole. We need to shift the bits 
        ' bits in each byte and sum them all together by multiplying by powers of 2 
        ' So the sub authority would be built by the following formula: 
        ' 
        ' SUbAuthority = byte0*2^0 + Byte1*2^8 + byte2*2^16 + byte3*2^24 
        ' 
        ' this be done using a simple short loop, initializing the power of two 
        ' variable ( p2 ) to 0 before the start an incrementing by 8 on each byte 
        ' and summing them all together. 
        ' 
        p2 = 0 
        subtotal = 0 
        for j = 1 to 4 
        dblSubAuth = CDbl(ascb(midb(osid,i,1))) * (2^p2) 
        subTotal = subTotal + dblSubAuth 
        p2 = p2 + 8 
        i = i + 1 
    next 
' 
' Convert the value to a string, add it to the SDDL Sid and continue 
' 
    strSDDL = strSDDL & "-" & cstr(subTotal) 
    next 
    SDDL_SID = strSDDL 
end function


'===Clear Variables==='
Set strComputer = Nothing 
Set strSourceAcct = Nothing 
Set strSourceAcctDomain = Nothing 
Set strTargetAcct = Nothing
Set strTargetAcctDomain = Nothing 
Set strTargetAcctSID = Nothing 
Set objComputer = Nothing 
Set objProfile = Nothing 
Set objCommand = Nothing 
Set objRecordSet = Nothing 
Set objConnection = Nothing 
Set objWMIService = Nothing 
Set objSID = Nothing 
Set dtStart = Nothing 
Set colProfiles = Nothing 
Set oSID = Nothing 
Set oUsr = Nothing 
Set strDomainDN = Nothing
Set Revision = Nothing 
Set IssueAuthorities(11) = Nothing 
Set strSDDL = Nothing 
Set subAuthorities = Nothing
Set strDomAcctLength = Nothing 
Set strCompName = Nothing 
Set strSlashPos = Nothing 
Set strContinue = Nothing
Set strDomainAcct = Nothing 
Set strLocalAcct = Nothing 
Set strLocalAcctLength = Nothing
Set strArg3 = Nothing
Set strKeepuser = Nothing 
Set strArg4 = Nothing
Set strWmios = Nothing
Set strOsQuery = Nothing
Set strOSVer = Nothing
Set colOperatingSystems = Nothing
Set objOs = Nothing
Set strSrcDom = Nothing


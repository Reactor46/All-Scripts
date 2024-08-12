'=*=*=*=*=*=*=*=*=*=*=*=*=
' Author : Assaf Miron 
' Http://assaf.miron.googlepages.com
' Date : 09/08/10
' GetUserSID.vbs
' Description : Gets the User SID and DN
'=*=*=*=*=*=*=*=*=*=*=*=*=
Option Explicit
On Error Resume Next
Function FindObject(ObjClass, SearchCret, strObj)
	Const ADS_SCOPE_SUBTREE = 2
	Dim objRootDSE,objConnection,objCommand,objRecordSet
	Dim strDomainLdap

	Set objRootDSE = GetObject ("LDAP://rootDSE")
	strDomainLdap  = objRootDSE.Get("defaultNamingContext")
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.CommandText = _
		"SELECT AdsPath FROM 'LDAP://" & strDomainLdap &_
			"' WHERE objectClass='" & ObjClass & "' and " & SearchCret & "='" &_
				strObj & "'"
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Timeout") = 30
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False
	
	Set objRecordSet = objCommand.Execute
	
	If objRecordSet.RecordCount = 0 Then 
		FindObject= 0
	Else 
		objRecordSet.Requery
		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			FindObject= objRecordSet.Fields("AdsPath").Value
			objRecordSet.MoveNext
		Loop
	End If 	
End Function 

UserName = "AssafM"
Set objUserName = GetObject(FindObject("user", "sAMAccountName", UserName))


Dim arrSid
Dim strSidHex, strSidDec, strObjectType

arrSid = objUserName.objectSid

strSidHex = OctetToHexStr(arrSid)

strSidDec = HexSIDtoSDDL(strSidHex)

WScript.Echo "User Name:" & objUserName.sAMAccountName
wscript.echo "Object DN: " & vbTab & objUserName.distinguishedName
Wscript.echo "Hex SID: " & vbTab & strSidHex
Wscript.echo "SDDL SID: " & vbTab & strSidDec

'Function to convert OctetString (byte array) to Hex string.
Function OctetToHexStr(arrbytOctet)
  Dim k
  OctetToHexStr = ""
  For k = 1 To Lenb(arrbytOctet)
    OctetToHexStr = OctetToHexStr & Right("0" & Hex(Ascb(Midb(arrbytOctet, k,1))), 2)
  Next
End Function

' Function to convert hex Sid to decimal (SDDL) Sid.
Function HexSIDtoSDDL(strHexSID)
  Dim i
  Dim strA, strB, strC, strD, strE, strF, strG

  ReDim arrTemp(Len(strHexSID)/2 - 1)

  'Create an array, where each element contains a single byte from the hex number
  For i = 0 To UBound(arrTemp)
    arrTemp(i) = Mid(strHexSID, 2 * i + 1, 2)
  Next

  'Move through the array to get each section, then convert it to decimal format

  strA = CInt(arrTemp(0))

  For i = 0 To UBound(arrTemp) 'Forward cycle for big-endian format
    Select Case i
      Case 2 strB = strB & arrTemp(i)
      Case 3 strB = strB & arrTemp(i)
      Case 4 strB = strB & arrTemp(i)
      Case 5 strB = strB & arrTemp(i)
      Case 6 strB = strB & arrTemp(i)
      Case 7 strB = strB & arrTemp(i)
    End Select
  Next
  strB = CInt("&H" & strB)

  For i = UBound(arrTemp) To 0 Step -1 'Reverse cycle for little-endian format
    Select Case i
      Case 11 strC = strC & arrTemp(i)
      Case 10 strC = strC & arrTemp(i)
      Case 9 strC = strC & arrTemp(i)
      Case 8 strC = strC & arrTemp(i)
    End Select
  Next
  strC = CInt("&H" & strC)

  For i = UBound(arrTemp) To 0 Step -1 'Reverse cycle for little-endian format
    Select Case i
      Case 15 strD = strD & arrTemp(i)
      Case 14 strD = strD & arrTemp(i)
      Case 13 strD = strD & arrTemp(i)
      Case 12 strD = strD & arrTemp(i)
    End Select
  Next
  strD = CLng("&H" & strD)

  For i = UBound(arrTemp) To 0 Step -1 'Reverse cycle for little-endian format
    Select Case i
      Case 19 strE = strE & arrTemp(i)
      Case 18 strE = strE & arrTemp(i)
      Case 17 strE = strE & arrTemp(i)
      Case 16 strE = strE & arrTemp(i)
    End Select
  Next
  strE = CLng("&H" & strE)

  For i = UBound(arrTemp) To 0 Step -1 'Reverse cycle for little-endian format
    Select Case i
      Case 23 strF = strF & arrTemp(i)
      Case 22 strF = strF & arrTemp(i)
      Case 21 strF = strF & arrTemp(i)
      Case 20 strF = strF & arrTemp(i)
    End Select
  Next
  strF = CLng("&H" & strF)

  For i = UBound(arrTemp) To 0 Step -1 'Reverse cycle for little-endian format
    Select Case i
      Case 27 strG = strG & arrTemp(i)
      Case 26 strG = strG & arrTemp(i)
      Case 25 strG = strG & arrTemp(i)
      Case 24 strG = strG & arrTemp(i)
    End Select
  Next
  strG = CLng("&H" & strG)

  HexSIDtoSDDL = "S-" & strA & "-" & strB & "-" & strC & "-" & strD & "-" & strE & "-" & strF & "-" & strG

End Function

Function HexStrToDecStr(strSid)
  ' Function to convert Hex string Sid to Decimal string (SDDL) Sid.

  ' SID anatomy:
  ' Byte Position
  ' 0 : SID Structure Revision Level (SRL)
  ' 1 : Number of Subauthority/Relative Identifier
  ' 2-7 : Identifier Authority Value (IAV) [48 bits]
  ' 8-x : Variable number of Subauthority or Relative Identifier (RID) [32bits]
  '
  ' Example:
  '
  ' <Domain/Machine>\Administrator
  ' Pos : 0 | 1 | 2 3 4 5 6 7 | 8 9 10 11 | 12 13 14 15 | 16 17 18 19 | 20 21 22 23 | 24 25 26 27
  ' Value: 01 | 05 | 00 00 00 00 00 05 | 15 00 00 00 | 06 4E 7D 7F | 11 57 56 7A | 04 11 C5 20 | F4 01 00 00
  ' str : S- 1 | | -5 | -21 | -2138918406 | -2052478737 | -549785860 | -500


  Const BYTES_IN_32BITS = 4
  Const SRL_BYTE = 0
  Const IAV_START_BYTE = 2
  Const IAV_END_BYTE = 7
  Const RID_START_BYTE = 8
  Const MSB = 3 'Most significant byte
  Const LSB = 0 'Least significant byte

  Dim arrbytSid, lngTemp, base, offset, i

  ReDim arrbytSid(Len(strSid)/2 - 1)

  ' Convert hex string into integer array
  For i = 0 To UBound(arrbytSid)
    arrbytSid(i) = CInt("&H" & Mid(strSid, 2 * i + 1, 2))
  Next

  ' Add SRL number
  HexStrToDecStr = "S-" & arrbytSid(SRL_BYTE)

  ' Add Identifier Authority Value
  lngTemp = 0
  For i = IAV_START_BYTE To IAV_END_BYTE
    lngTemp = lngTemp * 256 + arrbytSid(i)
  Next
  HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

  ' Add a variable number of 32-bit subauthority or
  ' relative identifier (RID) values.
  ' Bytes are in reverse significant order.
  ' i.e. HEX 01 02 03 04 => HEX 04 03 02 01
  ' = (((0 * 256 + 04) * 256 + 03) * 256 + 02) * 256 + 01
  ' = DEC 67305985
  For base = RID_START_BYTE To UBound(arrbytSid) Step BYTES_IN_32BITS
    lngTemp = 0
    For offset = MSB to LSB Step -1
      lngTemp = lngTemp * 256 + arrbytSid(base + offset)
    Next
    HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
  Next
End Function ' HexStrToDecStr
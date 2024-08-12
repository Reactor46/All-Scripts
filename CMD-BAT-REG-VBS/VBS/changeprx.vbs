' Script credits
' Sue Mosher http://www.outlookcode.com/codedetail_print.aspx?id=821
' Richard Mueller http://groups.google.com.au/group/microsoft.public.windows.server.active_directory/browse_thread/thread/5e26b20bba486280/95eb5a589be4cf11
' Vinay Pal Singh http://smarthost.blogspot.com/2009/03/outlook-anywhere-or-rpc-over-https.html
' Oz Casey Dedeal http://smtp25.blogspot.com/2009/03/rpc-over-https-script.html
'
SourceServer = wscript.arguments(2)
TargetServer = wscript.arguments(3)

strComputer = wscript.arguments(0)
strUser =  wscript.arguments(1)

set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(&(objectclass=person)(samaccountname=" & strUser & "));name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
 Userdn = rs.fields("distinguishedName")
 wscript.echo Userdn
 rs.movenext
Wend
Set objUser = GetObject("LDAP://" & Userdn)
usrsid =  ObjSidToStrSid(objUser.objectSid)

wscript.echo "UserSid : " &  usrsid 

Const HKEY_USERS = &H80000003


Set oReg =GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
strComputer & "\root\default:StdRegProv")

prProxyVal = "001f6622"
crcertVal = "001f6625"

strKeyPath = usrsid & "\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"

oReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
	strFullPath = strKeyPath & "\" & subkey & "\13dbb0c8aa05101a9bb000aa002fc45a"
	oReg.GetBinaryValue HKEY_USERS,strFullPath,prProxyVal,prProxyValRes
	sval = ""
	if not IsNull(prProxyValRes)  then 
		For Each byteValue in prProxyValRes
			sval = sval & ChrB(byteValue)
		Next
		if mid(sval,1,len(SourceServer)) = SourceServer then
			newarrValue = StringToByteArray(TargetServer, True)
			newarrValue2 = StringToByteArray(("msstd:"  & TargetServer), True)
			oReg.SetBinaryValue HKEY_USERS,strFullPath,prProxyVal,newarrValue
			oReg.SetBinaryValue HKEY_USERS,strFullPath,crcertVal,newarrValue2
			wscript.echo "Changed : " & sval 
		end if 
	end if

Next

Public Function StringToByteArray _
                 (Data, NeedNullTerminator)
    Dim strAll
    strAll = StringToHex4(Data)
    If NeedNullTerminator Then
        strAll = strAll & "0000"
    End If
    intLen = Len(strAll) \ 2
    ReDim arr(intLen - 1)
    For i = 1 To Len(strAll) \ 2
        arr(i - 1) = CByte _
                   ("&H" & Mid(strAll, (2 * i) - 1, 2))
    Next
    StringToByteArray = arr
End Function


Public Function StringToHex4(Data)
    ' Input: normal text
    ' Output: four-character string for each character,
    '         e.g. "3204" for lower-case Russian B,
    '        "6500" for ASCII e
    ' Output: correct characters
    ' needs to reverse order of bytes from 0432
    Dim strAll
    For i = 1 To Len(Data)
        ' get the four-character hex for each character
        strChar = Mid(Data, i, 1)
        strTemp = Right("00" & Hex(AscW(strChar)), 4)
        strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
    Next
    StringToHex4 = strAll
End Function

Function ArrayToMB(A)
  Dim I, MB
  For I = LBound(A) To UBound(A)
    MB = MB & ChrB(A(I))
  Next
  ArrayToMB = MB
End Function


Function ObjSidToStrSid(arrSid)
' Function to convert OctetString (byte array) to Decimal string (SDDL)
Dim strHex, strDec

strHex = OctetStrToHexStr(arrSid)
strDec = HexStrToDecStr(strHex)
ObjSidToStrSid = strDec
End Function ' ObjSidToStrSid

Function OctetStrToHexStr(arrbytOctet)
' Function to convert OctetString (byte array) to Hex string.
Dim k

OctetStrToHexStr = ""
For k = 1 To Lenb(arrbytOctet)
OctetStrToHexStr = OctetStrToHexStr _
& Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
Next
End Function ' OctetStrToHexStr

Function HexStrToDecStr(strSid)
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
For base = RID_START_BYTE To UBound(arrbytSid) Step BYTES_IN_32BITS
lngTemp = 0
For offset = MSB to LSB Step -1
lngTemp = lngTemp * 256 + arrbytSid(base + offset)
Next
HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
Next

End Function ' HexStrToDecStr




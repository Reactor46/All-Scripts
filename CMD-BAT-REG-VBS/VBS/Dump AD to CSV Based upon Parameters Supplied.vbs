Option Explicit

Dim adoCommand, adoConnection, strBase, strFilter, strAttributes
Dim objRootDSE, strBaseDN, strQuery, adoRecordset
Dim arrAttributes, k, intCount, strValue, strItem, strType
Dim objValue, lngHigh, lngLow, lngValue, strAttr, dtmValue
Dim objShell, lngBiasKey, lngBias, dtmDate, blnCSV, strLine
Dim strMulti, strArg

blnCSV = True
If (Wscript.Arguments.Count = 1) Then
    strArg = Wscript.Arguments(0)
    Select Case LCase(strArg)
        Case "/csv"
            blnCSV = True
    End Select
End If

' Obtain local Time Zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If
Set objShell = Nothing

' Setup ADO objects.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Prompt for base of query.
strBaseDN = Trim(InputBox("Specify DN of base of query, or blank for entire domain"))
If (strBaseDN = "") Then
    ' Search entire Active Directory domain.
    Set objRootDSE = GetObject("LDAP://RootDSE")
    strBaseDN = objRootDSE.Get("defaultNamingContext")
End If
If (InStr(LCase(strBaseDN), "dc=") = 0) Then
    Set objRootDSE = GetObject("LDAP://RootDSE")
    strBaseDN = strBaseDN & "," & objRootDSE.Get("defaultNamingContext")
    strBaseDN = Replace(strBaseDN, ",,", ",")
End If
strBase = "<LDAP://" & strBaseDN & ">"

' Prompt for filter.
strFilter = Trim(InputBox("Enter LDAP syntax filter"))
If (Left(strFilter, 1) <> "(") Then
    strFilter = "(" & strFilter
End If
If (Right(strFilter, 1) <> ")") Then
    strFilter = strFilter & ")"
End If

' Prompt for attributes.
strAttributes = InputBox("Enter comma delimited list of attribute values to retrieve")
strAttributes = Replace(strAttributes, " ", "")
strAttr = strAttributes
If (strAttributes = "") Then
    strAttributes = "distinguishedName"
Else
    strAttributes = "distinguishedName" & "," & strAttributes
End If
arrAttributes = Split(strAttributes, ",")

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

If (blnCSV = False) Then
    Wscript.Echo "Base of query: " & strBaseDN
    Wscript.Echo "Filter: " & strFilter
    Wscript.Echo "Attributes: " & strAttr
End If

' Run the query.
' Trap possible error.
On Error Resume Next
Set adoRecordset = adoCommand.Execute
If (Err.Number <> 0) Then
    Select Case Err.Number
        Case -2147217865
            Wscript.Echo "Table does not exist. Base of search not found."
        Case -2147217900
            Wscript.Echo "One or more errors. Filter syntax error."
        Case -2147467259
            Wscript.Echo "Unspecified error. Invalid attribute name."
        Case Else
            Wscript.Echo "Error: " & Err.Number
            Wscript.Echo "Description: " & Err.Description
    End Select
    Wscript.Quit
End If
On Error GoTo 0

' Enumerate the resulting recordset.
intCount = 0
Do Until adoRecordset.EOF
    ' Retrieve values and display.
    intCount = intCount + 1
    If (blnCSV = True) Then
        strLine = """" & adoRecordset.Fields("distinguishedName").Value & """"
    Else
        Wscript.Echo "DN: " & adoRecordset.Fields("distinguishedName").Value
    End If
    For k = 1 To UBound(arrAttributes)
        strType = TypeName(adoRecordset.Fields(arrAttributes(k)).Value)
        If (strType = "Object") Then
            Set objValue = adoRecordset.Fields(arrAttributes(k)).Value
            lngHigh = objValue.HighPart
            lngLow = objValue.LowPart
            If (lngLow < 0) Then
                lngHigh = lngHigh + 1
            End If
            lngValue = (lngHigh * (2 ^ 32)) + lngLow
            If (lngValue > 120000000000000000) Then
                dtmValue = #1/1/1601# + (lngValue/600000000 - lngBias)/1440
                On Error Resume Next
                dtmDate = CDate(dtmValue)
                If (Err.Number <> 0) Then
                    On Error GoTo 0
                    If (blnCSV = True) Then
                        strLine = StrLine & ",<Never>"
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & FormatNumber(lngValue, 0) _
                            & " <Never>"
                    End If
                Else
                    On Error GoTo 0
                    If (blnCSV = True) Then
                        strLine = strLine & "," & CStr(dtmDate)
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & FormatNumber(lngValue, 0) _
                            & " (" & CStr(dtmDate) & ")"
                    End If
                End If
            Else
                If (blnCSV = True) Then
                    strLine = strLine & ",""" & FormatNumber(lngValue, 0) & """"
                Else
                    Wscript.Echo "  " & arrAttributes(k) _
                        & ": " & FormatNumber(lngValue, 0)
                End If
            End If
        Else
            strValue = adoRecordset.Fields(arrAttributes(k)).Value
            Select Case strType
                Case "String"
                    If (blnCSV = True) Then
                        strLine = strLine & ",""" & strValue & """"
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & strValue
                    End If
                Case "Variant()"
                    strMulti = ""
                    For Each strItem In strValue
                        If (blnCSV = True) Then
                            If (strMulti = "") Then
                                strMulti = """" & strItem & """"
                            Else
                                strMulti = strMulti & ";""" & strItem & """"
                            End If
                        Else
                            Wscript.Echo "  " & arrAttributes(k) _
                                & ": " & strItem
                        End If
                    Next
                    If (blnCSV = True) Then
                        strLine = strLine & "," & strMulti
                    End If
                Case "Long"
                    If (blnCSV = True) Then
                        strLine = strLine & ",""" & FormatNumber(strValue, 0) & """"
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & FormatNumber(strValue, 0)
                    End If
                Case "Boolean"
                    If (blnCSV = True) Then
                        strLine = strLine & "," & CBool(strValue)
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & CBool(strValue)
                    End If
                Case "Date"
                    If (blnCSV = True) Then
                        strLine = strLine & "," & CDate(strValue)
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & CDate(strValue)
                    End If
                Case "Byte()"
                    If (blnCSV = True) Then
                        strLine = strLine & "," & OctetToHexStr(strValue)
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": " & OctetToHexStr(strValue)
                    End If
                Case "Null"
                    If (blnCSV = True) Then
                        strLine = strLine & ",<no value>"
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": <no value>"
                    End If
                Case Else
                    If (blnCSV = True) Then
                        strLine = strLine & ",<unsupported syntax>"
                    Else
                        Wscript.Echo "  " & arrAttributes(k) _
                            & ": <unsupported syntax " & TypeName(strValue) & " >"
                    End If
            End Select
        End If
    Next
    If (blnCSV = True) Then
        Wscript.Echo strLine
    End If
    adoRecordset.MoveNext
Loop
If (blnCSV = False) Then
    Wscript.Echo "Number of objects found: " & CStr(intCount)
End If

' Clean up.
adoRecordset.Close
adoConnection.Close

Function OctetToHexStr(ByVal arrbytOctet)
    ' Function to convert OctetString (byte array) to Hex string.

    Dim k

    OctetToHexStr = ""
    For k = 1 To Lenb(arrbytOctet)
        OctetToHexStr = OctetToHexStr _
            & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
    Next

End Function
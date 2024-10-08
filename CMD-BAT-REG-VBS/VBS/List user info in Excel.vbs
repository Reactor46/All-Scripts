Option Explicit
Const ADS_SCOPE_SUBTREE = 2
Sub LoadUserInfo()
    Dim x, objConnection, objCommand, objRecordSet, oUser, skip, disa
    Dim sht As Worksheet
    
    ' get domain
    Dim oRoot
    Set oRoot = GetObject("LDAP://rootDSE")
    Dim sDomain
    sDomain = oRoot.Get("defaultNamingContext")
    Dim strLDAP
    strLDAP = "LDAP://" & sDomain
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 100
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    
    objCommand.CommandText = "SELECT adsPath FROM '" & strLDAP & "' WHERE objectCategory='person'AND objectClass='user'"
    Set objRecordSet = objCommand.Execute
        
    x = 2
    Set sht = ThisWorkbook.Worksheets("Company")
    With sht
        ' Clear and set Header info
        .Cells.Clear
        .Cells.NumberFormat = "@"
        .Cells(1, 1).Value = "Login"
        .Cells(1, 2).Value = "Name"
        .Cells(1, 3).Value = "Surmane"
        .Cells(1, 4).Value = "Display Name"
        .Cells(1, 5).Value = "Departement"
        .Cells(1, 6).Value = "Title"
        .Cells(1, 7).Value = "Telephone"
        .Cells(1, 8).Value = "Mobile"
        .Cells(1, 9).Value = "Fax"
        .Cells(1, 10).Value = "Initials"
        .Cells(1, 11).Value = "Company"
        .Cells(1, 12).Value = "Address"
        .Cells(1, 13).Value = "P.O. box"
        .Cells(1, 14).Value = "Zip"
        .Cells(1, 15).Value = "Town"
        .Cells(1, 16).Value = "State"
        Do Until objRecordSet.EOF
            Set oUser = GetObject(objRecordSet.Fields("aDSPath"))
            skip = oUser.sAMAccountName
            disa = oUser.AccountDisabled
                        
            If (skip = "Administrator") Or (skip = "Guest") Or (skip = "krbtgt") Or (disa = "True") Then
                .Cells(x, 1).Value = "test"
                DoEvents
                objRecordSet.MoveNext
            Else
                .Cells(x, 1).Value = CStr(oUser.sAMAccountName) 'Replace(oUser.Name, "CN=", "")
                .Cells(x, 2).Value = oUser.givenName
                .Cells(x, 3).Value = oUser.SN
                .Cells(x, 4).Value = oUser.displayName
                .Cells(x, 5).Value = oUser.department
                .Cells(x, 6).Value = oUser.Title
                .Cells(x, 7).Value = oUser.telephoneNumber
                .Cells(x, 8).Value = oUser.mobile
                .Cells(x, 9).Value = oUser.facsimileTelephoneNumber
                .Cells(x, 10).Value = oUser.initials
                .Cells(x, 11).Value = oUser.company
                .Cells(x, 12).Value = oUser.streetAddress
                .Cells(x, 13).Value = oUser.postOfficeBox
                .Cells(x, 14).Value = oUser.postalCode
                .Cells(x, 15).Value = oUser.l ' by
                .Cells(x, 16).Value = oUser.st
                DoEvents
                x = x + 1
                objRecordSet.MoveNext
            End If
            
        Loop
        
    End With
    Range("A1:D1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("C12").Select

End Sub


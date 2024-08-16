Email = "smtp:" & wscript.arguments(1)
ExchangeServer = wscript.arguments(0)
Set rootDSE = GetObject("LDAP://RootDSE")
domainContainer =  rootDSE.Get("defaultNamingContext")
Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADSDSOObject"
conn.Open "ADs Provider"
LDAPStr = "<LDAP://" & DomainContainer &  ">;(&(objectCategory=publicfolder)(proxyAddresses=" & email & "));adspath,objectguid;subtree"
Set rs = conn.Execute(LDAPStr)
If rs.RecordCount = 1 Then
	wscript.echo FindPublicFolderWMI(transposeGuid(ConvertObjectGuidToString(rs.fields("objectguid"))))
End If

Function ConvertObjectGuidToString(ByVal arrRawObjectGUID)
Dim i, strByte
Dim arrObjectGUID(15)
For i = 1 To LenB(arrRawObjectGUID)
strByte = Hex(AscB(MidB(arrRawObjectGUID, i, 1)))
If Len(strByte) = 1 Then strByte = "0" & strByte
arrObjectGUID(i - 1) = strByte
Next
ConvertObjectGuidToString = Join(arrObjectGUID, "")
End Function

Function transposeGuid(guid) 
transposeGuid = "{" & mid(guid,7,2) & mid(guid,5,2) & mid(guid,3,2) _ 
 & mid(guid,1,2) & "-" & mid(guid,11,2) & mid(guid,9,2) _ 
 & "-" & mid(guid,15,2) & mid(guid,13,2) & "-" & mid(guid,17,4) _ 
        & "-" & mid(guid,21,12) & "}"
end function 

Function FindPublicFolderWMI(AdproxyPath)

Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_PublicFolder"
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//"& _
ExchangeServer &"/"&cWMINameSpace
Set objWMIServices = GetObject(strWinMgmts)
Set objPubInstances = objWMIServices.ExecQuery ("Select * From Exchange_PublicFolder Where adproxyPath='" & AdproxyPath & "'")
For Each objExchange_PublicFolder in objPubInstances
	path = objExchange_PublicFolder.Path
Next
FindPublicFolderWMI = path
End function
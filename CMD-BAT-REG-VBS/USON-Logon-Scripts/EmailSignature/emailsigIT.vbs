On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Title
strCred = objUser.info
strStreet = objUser.StreetAddress
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strFax = objUser.FacsimileTelephoneNumber
strEmail = objUser.mail
strCompany = objuser.company
strHomePage = objuser.HomePage
strPOBox = objUser.postOfficeBox
strDepartment = objuser.Department  

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries



objSelection.Font.Name = "Arial"
objSelection.Font.Size = 9
objSelection.Font.Bold = True
objSelection.Font.Color = RGB(212,96,39)
objSelection.TypeText "__________________________________________________________"
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)

objSelection.Font.Color = RGB (0,0,0)
objSelection.TypeText strName & " | " & strCompany
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.TypeText Chr(11)


objSelection.TypeText strTitle & " | " & strDepartment
objSelection.TypeText Chr(11)
objSelection.TypeText Chr(11)

if (strStreet) Then objSelection.TypeText strStreet & Chr(11)

if (strPOBox) Then objSelection.TypeText strPOBox & Chr(11)

if (strPhone) Then objSelection.TypeText "T     " & strPhone & Chr(11)

if (strFax) Then objSelection.TypeText "F     " & strFax & Chr(11)

if (strMobile) Then objSelection.TypeText "M     " & strMobile & Chr(11)


objSelection.Font.Color = RGB(49,61,177)
objSelection.TypeText strEmail 
objSelection.TypeText Chr(11)

objSelection.TypeText strHomePage
objSelection.TypeText Chr(11)

objSelection.Font.Color = RGB(212,96,39)
objSelection.TypeText "__________________________________________________________"
objSelection.TypeText Chr(11)

objSelection.Font.Color = Black

Set objSelection = objDoc.Range()


objSignatureEntries.Add "Standard Signature", objSelection
objSignatureObject.NewMessageSignature = "Standard Signature"
objSignatureObject.ReplyMessageSignature = "Standard Signature"


objDoc.Saved = True
objWord.Quit


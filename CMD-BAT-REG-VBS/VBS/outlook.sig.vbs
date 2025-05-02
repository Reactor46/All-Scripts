On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Description
strCred = objUser.info
strStreet = objUser.StreetAddress
strLocation = objUser.l
strPostCode = objUser.PostalCode
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strFax = objUser.FacsimileTelephoneNumber
strEmail = objUser.mail

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName
objSelection.TypeParagraph()
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)
objselection.TypeText Chr(11)
objSelection.TypeText "Company Name"
objSelection.TypeText Chr(11)
objSelection.TypeText strStreet
objSelection.TypeText Chr(11)
objSelection.TypeText "PHONE: " & strPhone
objSelection.TypeText Chr(11)
if (strFax) Then objSelection.TypeText "FAX: " & strFax & Chr(11)
if (strMobile) Then objSelection.TypeText "CELL: " & strMobile & Chr(11)
objSelection.TypeText "EMAIL: " & strEmail
objSelection.TypeText Chr(11)
objSelection.TypeText "________________________________"
objSelection.TypeText Chr(11)
objSelection.TypeText "CONFIDENTIALITY NOTICE"

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Full Signature", objSelection
objSignatureObject.NewMessageSignature = "Full Signature"

objDoc.Saved = True
objWord.Quit

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName
objSelection.TypeParagraph()
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Reply Signature", objSelection

objSignatureObject.ReplyMessageSignature = "Reply Signature"

objDoc.Saved = True
objWord.Quit
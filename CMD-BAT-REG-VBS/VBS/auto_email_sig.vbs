'Auto Create Email Siganture 

On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Title
strFax = objUser.facsimileTelephoneNumber
strCompany = objUser.Company
strPhone = objUser.telephoneNumber

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

objSelection.Font.Name = "Calibri"
objSelection.Font.Bold = True
objSelection.Font.Size = "12"
objSelection.Font.Color = RGB(102,158,0)
objSelection.TypeText strName
objSelection.TypeParagraph()
objSelection.Font.Name = "Calibri"
objSelection.Font.Bold = False
objSelection.Font.Size = "11"
objSelection.Font.Color = RGB(105,105,105)
objSelection.TypeText strTitle
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.Font.Name = "Calibri"
objSelection.Font.Bold = True
objSelection.Font.Size = "12"
objSelection.Font.Color = RGB(105,105,105)
objSelection.TypeText strCompany
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.Font.Name = "Calibri"
objSelection.Font.Bold = False
objSelection.Font.Underline = True
objSelection.Font.Size = "11"
objSelection.Font.Color = RGB(102,158,0)
objSelection.TypeText "www.yourURL.com"
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.Font.Name = "Calibri"
objSelection.Font.Bold = False
objSelection.Font.Underline = False
objSelection.Font.Size = "11"
objSelection.Font.Color = RGB(105,105,105)
objSelection.TypeText "555.555.5555 "
objSelection.Font.Color = RGB(102,158,0)
objSelection.TypeText "Office"
objSelection.TypeParagraph()
objSelection.Font.Color = RGB(105,105,105)
objSelection.TypeText strPhone
objSelection.Font.Color = RGB(102,158,0)
objSelection.TypeText " Direct"
objSelection.TypeParagraph()
objSelection.Font.Color = RGB(105,105,105)
objSelection.TypeText strFax
objSelection.Font.Color = RGB(102,158,0)
objSelection.TypeText " Fax"
objSelection.TypeParagraph()


Set objSelection = objDoc.Range()

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"

objDoc.Saved = True
objWord.Quit

'End Emaail Signature
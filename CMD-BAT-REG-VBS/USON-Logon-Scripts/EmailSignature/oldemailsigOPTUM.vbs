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

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries


objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10
objSelection.Font.Bold = True
objSelection.Font.Color = RGB (0,0,0)
objSelection.TypeText strName
objSelection.Font.Size = 8
objSelection.TypeText Chr(11)
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(212,96,39)
objSelection.TypeText strStreet
objSelection.TypeText Chr(11)
objSelection.TypeText "T: " & strPhone
if (strFax) Then objSelection.TypeText " | F: " & strFax
if (strMobile) Then objSelection.TypeText " | C: " & strMobile
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(49,61,177)
objSelection.TypeText strEmail & " | " & strHomePage
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture("\\uson.local\NETLOGON\EmailSignature\Optum_Email.bmp")
objSelection.TypeText Chr(11)
objSelection.TypeText "_______________________________________________________________"
objSelection.TypeText Chr(11)
objSelection.Font.Color = Black
objSelection.Font.Bold = True
objSelection.TypeText "CONFIDENTIALITY NOTICE:"
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(127,127,127)
objSelection.Font.Bold = False
objSelection.TypeText "The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain confidential and privileged information protected by privacy laws. If you are not the intended recipient, be advised that you have received this message in error and that any use, dissemination, forwarding, printing, or copying of this message or any of its attachments in any form is strictly prohibited. Please  destroy this message  and any attachments to this message. Your cooperation is greatly appreciated."

Set objSelection = objDoc.Range()


objSignatureEntries.Add "Standard Signature", objSelection
objSignatureObject.NewMessageSignature = "Standard Signature"
objSignatureObject.ReplyMessageSignature = "Standard Signature"


objDoc.Saved = True
objWord.Quit


On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Title
strCred = objUser.info
strStreet = "3150 N. Tenaya Way, Suite #160, Las Vegas, NV 89128"
strPhone = "702-724-8800"
strMobile = objUser.Mobile
strFax = "702-724-8801"
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
objSelection.Font.Color = RGB(79,129,189)
objSelection.TypeText strName
objSelection.Font.Size = 8
objSelection.TypeText Chr(11)
objSelection.TypeText strDescription
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(127,127,127)
objSelection.TypeText strStreet
objSelection.TypeText Chr(11)
objSelection.TypeText "T: " & strPhone
if (strFax) Then objSelection.TypeText " | F: " & strFax
if (strMobile) Then objSelection.TypeText " | C: " & strMobile
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(49,61,177)
objSelection.TypeText strEmail & " | " & strHomePage
objSelection.TypeText Chr(11)
objSelection.InlineShapes.AddPicture("\\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\ROsigimage.jpg")
objSelection.TypeText Chr(11)
objSelection.TypeText "_______________________________________________________________"
objSelection.TypeText Chr(11)
objSelection.Font.Color = Black
objSelection.Font.Bold = True
objSelection.TypeText "CONFIDENTIALITY NOTICE:"
objSelection.TypeText Chr(11)
objSelection.Font.Color = RGB(127,127,127)
objSelection.Font.Bold = False
objSelection.TypeText "The information contained in this electronic message and any attachments to this message are intended for the exclusive use of the addressee(s) and may contain confidential and privileged information protected by privacy laws. If you are not the intended recipient, be advised that you have received this message in error and that any use, dissemination, forwarding, printing, or copying of this message or any of its attachments in any form is strictly prohibited. Please destroy this message  and any attachments to this message. Your cooperation is greatly appreciated."

Set objSelection = objDoc.Range()


objSignatureEntries.Add "NVCS Signature", objSelection



objDoc.Saved = True
objWord.Quit


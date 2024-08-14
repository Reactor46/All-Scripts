'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------

' Change Compression Settings in PowerPoint 2007
' Target output:
' Print (220 ppi): excellent quality on most printers and screens
' Screen (150 ppi): good for Web pages and projectors
' E-mail (96 ppi): minimize document size for sharing

' Set reference to PowerPoint application.
Dim pptApp
' Set reference to Registry Object.
Dim regOperator

On Error Resume Next

' /* Create two objects. */
Set pptApp = CreateObject("PowerPoint.Application")
Set regOperator = WScript.CreateObject("WScript.Shell")

' /* If no error. */
If Err.Number = 0 Then
	' Define a variant to save the version of PowerPoint Application.
	Dim pptVer

	' Get the version of PowerPoint Application.
	pptVer = pptApp.Version

	Dim strRegKeyPath
	Dim strRegKeyTypeName

	' /* If the target machine has PowerPoint 2007 installed. */
	If pptVer = 12 Then

		strRegKeyPath = "HKCU\Software\Microsoft\Office\Common\CompressPictures\CompressDPISetting"
		strRegKeyTypeName = "REG_DWORD"

		Dim strInputVal

		strInputVal = InputBox ("Please enter a number (1 to 3):" & String(2, Chr(10)) _
						 & "Target output:" & Chr(10) _
						 & "1. Print (220 ppi)" & Chr(10) _
						 & "2. Screen (150 ppi)" & Chr(10) _
						 & "3. E-mail (96 ppi)" & Chr(10), "Input Required", 2)

		' /* If the input value is not empty. */
		If Not IsEmpty(strInputVal) Then
			If IsNumeric(strInputVal) Then
				If (strInputVal >= 1) And (strInputVal <= 3) Then
				
				Dim strRealVal
				
				strRealVal = Int(strInputVal) - 1
				
					' /* If "strRegKeyPath" already exists. */
					If regOperator.RegRead(strRegKeyPath) <> strRealVal Then
						regOperator.RegWrite strRegKeyPath, strRealVal, strRegKeyTypeName
						WScript.Echo "Set successfully!"

					' /* If "strRegKeyPath" not found. */
					Else
						WScript.Echo "Compression settings are not changed!"
					End If
				Else
					WScript.Echo "Input error!"
				End If
			Else
				WScript.Echo "Input error!"
			End If
		End If

	Else
		WScript.Echo "Application - PowerPoint 2007 cannot be found!"
	End If

Else
	WScript.Echo "Set failed!"
End If

' /* Release memory. */
Set pptApp = Nothing
Set regOperator = Nothing
' GLOABAL CONSTANTS
Const FOR_READING = 1
Dim g_StrPlainText
   g_StrPlainText = "plaintext.txt"
Dim g_StrChiperText
   g_StrChiperText = ".\chipertext.txt"

' Open File for Reading
Dim objFSOin
Dim objInputStream

Set objFSOin = CreateObject("Scripting.FileSystemObject")
If objFSOin.FileExists( g_StrPlainText ) Then
   Set objInputStream = objFSOin.OpenTextFile( g_StrPlainText, FOR_READING )
Else
   WScript.Echo "Input file " & g_StrPlainText & " not found."
   WScript.Quit
End If

' Create the output file
Dim objFSOout
Set objFSOout = CreateObject("Scripting.FileSystemObject")
Dim objOutputStream
set objOutputStream = objFSOout.CreateTextFile( g_StrChiperText )
'A counter for when to place a space in the chipher text
Dim spaceCount
spaceCount = 0
'A counter for when to place a carrage return in the chipher text
Dim lineCount
lineCount = 0

Dim strInput
Dim ansiInt

Do Until objInputStream.AtEndOfStream
   strInput = objInputStream.Read(1)
   ansiInt = Asc( strInput )

   ' Convert upper case 
   If ansiInt >= 97 Then                    ' 97  = a
      If ansiInt <= 122 Then                ' 122 = z
         ansiInt = ansiInt - 32             ' Subtracting 32 converts a to A
      End If
   End If

   If ansiInt <= 90 Then                              ' 90 = Z
      If ansiInt >= 65 Then                           ' 65 = A
         ansiInt = ( 7 * ansiInt + 10 ) Mod 26 + 65
         spaceCount = spaceCount + 1

         objOutputStream.Write( Chr( ansiInt ) )
      End If
   End If

   ' Add a space to the output when spaceCount equals 5
   If spaceCount = 5 Then
      ' Write the space
      objOutputStream.Write( " " )
      ' Reset spaceCount
      spaceCount = 0
      ' Increment lineCount
      lineCount = lineCount + 1
   End If

   ' Add a newline to the output when lineCount equals 10
   If lineCount = 10 Then
      objOutputStream.WriteLine()
      lineCount = 0
   End If
Loop

WScript.echo "Complete."

WScript.Quit

' A65 Z90 a97 z122
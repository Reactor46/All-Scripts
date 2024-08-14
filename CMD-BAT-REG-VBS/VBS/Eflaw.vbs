' Eflaw.vbs
' VBScript program to calculate the McAlpine EFLAW(TM) Readability Score.
' McAlpine EFLAW(TM) Readability Score - Copyright (c) 2006 Rachel McAlpine
' Based on "From Plain English to Global English", by Rachel McAlpine
' http://www.webpagecontent.com/arc_archive/139/5/
' Script Author: Richard L. Mueller
' Version 1.0 - October 3, 2012

Option Explicit

Dim strFile, objFSO, objFile, strText, intWords, intMini, objRE
Dim objWords, objWord, intSentences, objTerm, dblEflaw, strScore

Const ForReading = 1

' Check for file or prompt.
If (Wscript.Arguments.Count = 0) Then
    strFile = InputBox("Enter file name")
    ' Check if user canceled.
    If (strFile = "") Then
        Wscript.Quit
    End If
Else
    strFile = Wscript.Arguments(0)
End If

Wscript.Echo "McAlpine EFLAW(TM) Readability Score - Copyright (c) 2006 Rachel McAlpine"
Wscript.Echo "Script by Richard L. Mueller, Copyright (c) 2012, Version 1.0"
Wscript.Echo "----------"
Wscript.Echo "File: " & strFile

' Open the file.
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set objFile = objFSO.OpenTextFile(strFile, ForReading)
If (Err.Number <> 0) Then
    Wscript.Echo "Unable to open file: " & strFile
    Wscript.Echo "Error Number: " & Err.Number
    Wscript.Echo "Description: " & Err.Description
    Wscript.Quit
End If
On Error GoTo 0

' Read the file.
strText = objFile.ReadAll
objFile.Close

' Add terminating space, to make sure last sentence recognized.
strText = strText & " "

' Ignore characters, so they don't add to word lengths.
strText = Replace(strText, "'", "")
strText = Replace(strText, """", "")
strText = Replace(strText, "(", "")
strText = Replace(strText, ")", "")
strText = Replace(strText, ";", "")

' Replace dash and colon with space, so they delimit words.
strText = Replace(strText, ":", " ")
strText = Replace(strText, "-", " ")

Set objRE = New RegExp
objRE.Global = True

' Parse into words (no leading digits).
objRE.Pattern = "[a-zA-Z]\w*"
Set objWords = objRE.Execute(strText)
intWords = objWords.Count

' Count number of mini-words (three characters or less).
intMini =  0
For Each objWord In objWords
    If (objWord.Length < 4) Then
        intMini = intMini + 1
    End If
Next

' Count sentences (terminated by ".", "?", or "!"). This assumes there
' is no space before the ".", "?", or "!" that terminates a sentence.
objRE.Pattern = "\w*[\.|\?|\!][^0-9]"
Set objTerm = objRE.Execute(strText)
intSentences = objTerm.Count

' Calculate the McAlpine EFLAW(TM) Readability Score.
dblEflaw = (intWords + intMini)/intSentences
If (dblEflaw <= 20.49) Then
    strScore = "very easy to understand"
ElseIf (dblEflaw <= 25.49) Then
    strScore = "quite easy to understand"
ElseIf (dblEflaw <= 29.49) Then
    strScore = "a little difficult"
Else
    strScore = "very confusing"
End If

Wscript.Echo "Number of words: " & CStr(intWords)
Wscript.Echo "Number of mini-words: " & CStr(intMini)
Wscript.Echo "Number of sentences: " & CStr(intSentences)
Wscript.Echo "Words per sentence: " & FormatNumber((intWords/intSentences), 1)
Wscript.Echo "McAlpine EFLAW(TM) Readability Score: " & FormatNumber(dblEflaw, 1) _
    & " (" & strScore & ")"

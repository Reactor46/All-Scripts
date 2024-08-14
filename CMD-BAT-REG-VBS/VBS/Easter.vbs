'############################ Easter.vbs ##############################
'                          from h.r.roesler
' Computes the Easter date for a certain year or for a range of years
' (e.g. 2000ñ2012). If no year is specified, the date for the actual
' year is displayed.
' In this version all movable Christian holidays dependent from the
' Easter appointment are displayed in addition.
' Syntax:
' CScript.exe [path]Easter.vbs [1st Annual Number] [2nd Annual Number]
' Example:
' CScript.exe Easter.vbs 1950 1990
Option Explicit

Function Easter(X)   ' Formula from Gauﬂ, Lichtenberg et al
    ' from <http://www.ptb.de/cms/en/fachabteilungen/abt4/fb-44/ag-441/realisation-of-legal-time-in-germany/the-date-of-easter.html>
    Dim K, M, S, A, D, R, OG, SZ, OE

    K  = X \ 100
    M  = 15 + (3 * K + 3) \ 4 - (8 * K + 13) \ 25
    S  = 2 - (3 * K + 3) \ 4
    A  = X Mod 19
    D  = (19 * A + M) Mod 30
    R  = D \ 29 + (D \ 28 - D \ 29) * (A \ 11)
    OG = 21 + D - R
    SZ = 7 - (X + X \ 4 + S) Mod 7
    OE = 7 - (OG - SZ) Mod 7

    Easter = DateSerial(X, 3, OG + OE)
End Function

Sub DisplayDates(ByRef arrYear)
    Dim arr, i, OS, j, dat, str, f

    WSH.Echo vbCRLF & "The movable Christian celebrations fall on " & _
                      "these days:" & vbCRLF
    arr = Array(-46, "Ash Wednesday", 0, "Easter", 39, "Ascension", _
                49, "Whitsun", 60, "Corpus Christ")
    For i = arrYear(0) To arrYear(1) Step arrYear(2)
        If i > 1600 And i <= 9999 Then
            OS = Easter(i)
            For j = 0 To UBound(arr) Step 2
                dat = DateAdd("d", OS, arr(j))
                dat = FormatDateTime(dat, vbLongDate)
                str = str & arr(j + 1) & " on " & dat & vbCRLF
            Next
        ElseIf Not(f) Then
            WSH.StdErr.WriteLine "Years before 1601 and after 9999 in " & _
                      "Christian chronology cannot be considered." & vbCRLF
            f = True
        End If
        If Len(str) > 0 Then WSH.Echo str: str = ""
    Next
End Sub

Function EvalArgs(ByRef wshArgs)
    Dim str, i, x(2): Const UpLim = 1, LoLim = 0

    For Each str In wshArgs
        If IsNumeric(str) Then
            For i = LoLim To UBound(x) - UpLim
                If Len(Abs(str)) = 4 And IsEmpty(x(i)) Then
                    x(i) = Abs(str) \ 1: Exit For
                End If
            Next
        End If
        If IsEmpty(x(i)) Then WSH.Echo "Invalid Argument: " & str
    Next

    If IsEmpty(x(LoLim)) Then
        WSH.Echo vbCRLF & "Enter an annual number, and the likely " & _
                 "Easter date is computed." & vbCRLF & "For years " & _
                 "in the future no guarantee is granted particularly" & _
                  vbCRLF & "whether the forecasted date will also be" & _
                  " occur really."
        x(LoLim) = Year(Date)
    End If
    If IsEmpty(x(UpLim)) Then x(UpLim) = x(LoLim)
    If x(UpLim) < x(LoLim) Then x(2) = -1: Else x(2) = 1

    EvalArgs = x
End Function

Call DisplayDates(EvalArgs(WScript.Arguments))
'############################ Easter.vbs ##############################

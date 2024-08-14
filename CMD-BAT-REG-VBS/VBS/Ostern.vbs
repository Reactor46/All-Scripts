'############################ Ostern.vbs ##############################
' Berechnet das Osterdatum für ein bestimmtes Jahr oder für einen
' Bereich von Jahren (2000 - 2012 zum Beispiel). Wird kein Jahr
' angegeben, so wird das Datum für das aktuelle Jahr angezeigt.
' In dieser Version werden zusätzlich alle vom Ostertermin abhängigen
' beweglichen christlichen Festtage ausgegeben.
' Aufruf:
' CScript.exe [Pfad]Ostern.vbs [1. Jahreszahl] [2. Jahreszahl]
' Beispiel:
' CScript.exe Ostern.vbs 1950 1990
Option Explicit

Function Ostern(X)   ' Osterformel von Gauß, Lichtenberg und anderen
    ' nach <http://www.ptb.de/cms/fachabteilungen/abt4/fb-44/ag-441/darstellung-der-gesetzlichen-zeit/wann-ist-ostern.html>
    Dim K, M, S, A, D, R, OG, SZ, OE

    K  = X \ 100                     ' Säkularzahl
    M  = 15 + (3 * K + 3) \ 4 - (8 * K + 13) \ 25
    S  = 2 - (3 * K + 3) \ 4         ' säkulare Sonnenschaltung
    A  = X Mod 19                    ' Mondparameter
    D  = (19 * A + M) Mod 30         ' Keim 1. Vollmond im Frühling
    R  = D \ 29 + (D \ 28 - D \ 29) * (A \ 11) ' kalendarische Korrektur
    OG = 21 + D - R                 ' 1. Vollmond im Frühling
    SZ = 7 - (X + X \ 4 + S) Mod 7  ' 1. Sonntag im März
    OE = 7 - (OG - SZ) Mod 7

    Ostern = DateSerial(X, 3, OG + OE)
End Function

Sub DisplayDates(ByRef arrYear)
    Dim arr, i, OS, j, dat, str, f

    str = vbCRLF & "Die beweglichen christlichen Feste " & _
                      "fallen auf diese Tage:" & vbCRLF
    arr = Array(-46, "Aschermittwoch", 0, "Ostern", 39, "Vatertag", _
                49, "Pfingsten", 60, "Fronleichnam")
    For i = arrYear(0) To arrYear(1) Step arrYear(2)
        If i > 1600 And i <= 9999 Then
            OS = Ostern(i)
            For j = 0 To UBound(arr) Step 2
                dat = DateAdd("d", OS, arr(j))
                dat = FormatDateTime(dat, vbLongDate)
                str = str & arr(j + 1) & " auf " & dat & vbCRLF
            Next
        ElseIf Not(f) Then
            str = str & "Jahre vor 1601 und nach 9999 christlicher Zeitr" & _
                  "echnung können nicht berücksichtigt werden." & vbCRLF
            f = True
        End If
        If Len(str) > 0 Then WSH.Echo str: str = ""
    Next
End Sub

Function EvalArgs(ByRef wshArgs)
    Const UpLim = 1, LoLim = 0: Dim str, i, x(2)

    For Each str In wshArgs
        If IsNumeric(str) Then
            For i = LoLim To UBound(x) - UpLim
                If Len(Abs(str)) = 4 And IsEmpty(x(i)) Then
                    x(i) = Abs(str) \ 1: Exit For
                End If
            Next
        End If
        If IsEmpty(x(i)) Then WSH.Echo "Falsches Argument: " & str
    Next

    If IsEmpty(x(0)) Then
        WSH.Echo vbCRLF & "Gib eine Jahreszahl ein, und das wahr" & _
                 "scheinliche Osterdatum wird berechnet." & vbCRLF & _
                 "Für Jahre in der Zukunft wird keine Gewähr über" & _
                 "nommen, ob das prognosti-" & vbCRLF & "zierte " & _
                 "Datum auch tatsächlich eintreten wird." & vbCRLF
        x(0) = Year(Date)
    End If
    If IsEmpty(x(1)) Then x(1) = x(0)
    If x(1) < x(0) Then x(2) = -1: Else x(2) = 1

    EvalArgs = x
End Function

Call DisplayDates(EvalArgs(WScript.Arguments))
'############################ Ostern.vbs ##############################

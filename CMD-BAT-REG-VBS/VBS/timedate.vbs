On Error Resume Next
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("TimeDate.txt", True)
Set objTextFile = objFSO.OpenTextFile("computers.txt", ForReading)
Do Until objTextFile.AtEndOfStream 
    strComputer = objTextFile.Readline



    Set wmi = GetObject("winmgmts://" & strComputer & "/root/cimv2")

    Set op = wmi.ExecQuery("Select * From Win32_LocalTime")
    Dim Day, DayOfWeek, Month, Hour, Milli, Minute, Second, WeekInMonth, Year
     
    For Each ele In op
     Day = ele.Day
     Hour = ele.Hour
     Month = ele.Month
     Minute = ele.Minute
     Year = ele.Year
    Next
     
    Dim AmPm
   
    If Hour > 12 Then
     AmPm = "PM"
     Hour = Abs(Abs(Hour - 24) - 12)
    Else
     If Hour = 0 Then
         Hour = 12
     End If
        AmPm = "AM"
    End If
    If Minute < 10 Then
        Zero = 0
    Else
        Zero = ""
    End If  
objFile.WriteLine  " Source Time/Date is: " & Now & "    Current time at " & strComputer & " is " & _
                Month & "/" & Day & "/" & Year & _
             " " & Hour & ":" & Zero & Minute & " " & AmPm & vbCrLfv 

Loop
objTextFile.Close


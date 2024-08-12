'=*=*=*=*=*=*=*=*=*=*=*=
' Check if Computer List is Alive.vbs
' Coded By Assaf Miron 
' Date : 18/09/08
' Thanks to Assaf Israel
'=*=*=*=*=*=*=*=*=*=*=*=

Const ForReading = 1
Const ForAppending = 8

Dim Script_Name
Dim objFSO,objFile,objOutFile
Dim WshShell
Dim outLogFile
Dim arrComputersList
Dim strText,strComputer,strAlive
Dim ComputersFile,SleepTime

Script_Name = WScript.ScriptName

Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set WshShell = WScript.CreateObject("WScript.Shell")

Function IsAlive(strComputer)
'====================================================================
' Checks if the givven computer name Replies to Ping or not.
' If it does then it returns "Alive" Else the Computer must be Dead
'====================================================================
Dim objExecObject

Set objExecObject = WSHShell.Exec _
    ("%comspec% /c ping -n 3 -w 1000 " & strComputer)

' Error Running Ping
If objExecObject.ExitCode <> 0 Then
	IsAlive = "UnResolved"
Else
' Check if Computer is Alive
	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadAll()
	    If Instr(strText, "Reply") > 0 Then
	        IsAlive =  "Alive"
	    Else
	        IsAlive = "Dead"
	    End If 
	Loop
End If
End Function

Sub Usage
'====================================================================
' Show the Help Usage for the Script.
'====================================================================

WScript.Echo Replace(Script_Name,".vbs","") & " by Assaf Miron" & vbNewLine _
    & "Recieves a List of Computer and checks if they are Alive." & vbNewLine _ 
    & "Script Outputs to a CSV File, Located in the Input File Folder." & vbNewLine _
    & vbNewLine _
    & "Usage: " & SCRIPT_NAME & " [/Computers:[<Path>]] [/SleepTime <Num in MS>]" & vbNewLine _
    & vbNewLine _
    & "/Computers :" & vbNewLine & vbTab _
    & "<Path> : Path to a TXT/CSV File, this File will contain List of Computers" & vbNewLine & vbTab _
    & "Seperated by a new Line." & vbNewLine _
    & vbNewLine _
    & "/SleepTime :" & vbNewLine & vbTab _
    & "<Num in MS> : Number in MiliSeconds, Time to Sleep between Ping Commands." & vbNewLine & vbTab _
    & "If no Time is specified, the default Time is used (500ms)." & vbNewLine _
    & vbNewLine _
    & "Remarks :"

  Wscript.Quit
End Sub


' Main Code
' Check the Wscript Arguments 
' If no Arguments assinged then Show the Help Usage
If WScript.Arguments.Count = 0 Then
	Usage
Else
  With Wscript.Arguments
    If .Named.Exists("?") Then Usage
    If .Named.Exists("Computers") Then ComputersFile = .Named("Computers")
    If .Named.Exists("SleepTime") Then 
    	SleepTime = .Named("SleepTime")
	Else	
		SleepTime = 500
	End If    	
  End With
End If

' Checks if the File Exists - Prevent Problems
If Not objFSO.FileExists(ComputersFile) Then
	WScript.Echo "You Need to enter a Computers File." & vbNewLine _
		& "See Script Usage for Help (" & Script_Name & " /?)"
	
	WScript.Quit
End If

' Create the Output file in the same folder and name as the input File.
outLogFile = Mid(ComputersFile,1,Len(ComputersFile) -4) & "-Output.csv"
Set objFile = objFSO.OpenTextFile(ComputersFile, ForReading)
Set objOutFile = objFSO.CreateTextFile(outLogFile,ForAppending)


' Write CSV Headers for CSV File
objOutFile.WriteLine "Computer Name,Is Alive?"

' Read all the Text into strText
strText = objFile.ReadAll

If InStr(strText,",") Then ' File is Probably CSV
	arrComputersList = Split(strText , ",")	
Else
	arrComputersList = Split(strText , vbNewLine)
End If

For i = 0 to Ubound(arrComputersList)
	' Read Computer Name
	strComputer = arrComputersList(i) 
	' Check if Computer Alive
	strAlive = IsAlive(strComputer) 
	' Output Result
	objOutFile.WriteLine strComputer & "," & strAlive
	' Wait Until Next Run
	WScript.Sleep SleepTime
Next

' Done !
WScript.Echo "Script is Done!" & vbNewLine _
		   & " Output File Location : " & outLogFile
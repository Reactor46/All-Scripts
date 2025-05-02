'This Script will allow you to synch your domain, run a dcdiag or a netdiag 
'and dump append the results to a text file. This script is only intended to run 
'Against the local Server.


Option Explicit

'Will contain the user's response to the option list.
Dim myResponse
Dim myExit, WShell,myFSO,myTextStream,DCDIAG
Set WShell = CreateObject("WScript.Shell")
Set myFSO =  CreateObject("Scripting.FileSystemObject")

'Change this string for the DCDIAG command.
'Add /a after DCDIAG to make it run against all servers in enterprise.
'/ferr instead of /f only prints errors.
DCDIAG = "DCDIAG /f:c:\Diag.txt"
myExit = 0



Do While myExit = 0
'Get response from user.
myResponse = InputBox("What do you want to do? (S)ynch, (D)cdiag, (N)etdiag, (A)ll or (E)xit and launch log.", "DC Toolbox Menu")


'If response was s then run synch and dump to Diag.txt
If (myResponse = "S") or (myResponse = "s") or (myResponse = "a")Then
	Set myTextStream = myFSO.CreateTextFile("C:\temp.bat")
	myTextStream.Write "repadmin /syncall >>Diag.txt"
	myTextStream.Close
	WShell.Run "c:\temp.bat", , True
	myFSO.DeleteFile "C:\temp.bat"

End If

'If response was D then run DCDIAG and dump to Diag.txt
If (myResponse = "D") or (myResponse = "d") or (myResponse = "a")Then

	WShell.Run DCDIAG, , True

End If

'If response was N then run NetDiag and dump to Diag.txt
If (myResponse = "N") or (myResponse = "n") or (myResponse = "a")Then
	Set myTextStream = myFSO.CreateTextFile("C:\temp.bat")
	myTextStream.Write "Netdiag >>Diag.txt"
	myTextStream.Close
	WShell.Run "c:\temp.bat", , True
	myFSO.DeleteFile "C:\temp.bat"
End If

'If response was E then Exit.
If (myResponse = "E") or (myResponse = "e") or (myResponse = "") or (myResponse = "a")Then
	myExit = 1
	WShell.Run "Diag.txt"
End If

'End of loop
Loop
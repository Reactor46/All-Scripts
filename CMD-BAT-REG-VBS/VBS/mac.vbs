'-----------------------------------------------------------------------------------------------
' mac.vbs
' Author: Jeff Mason aka TNJMAN aka bitdoctor
' Date: 06/12/2011
' Purpose: Potentially, to be run at login for every user/computer, obtain MAC address
'  and write Computer Name, MAC address, IP Address(es) & Date to a text file (APPEND mode)
'  Output file will be in the format: c:\scripts\macinfo\computername-MAC.txt
'    where "computername" is the name of any computer against which this script was run
'  This script also is very useful for later running WOL tools to "wake computers," because MAC is 
'  needed for WOL, And incredibly useful for Auditing, Inventory & Security since, for each computer,
'  you can track WHEN a MAC or IP address changed (i.e. new/different NIC, spoofed NIC, etc.)
'
'
' To use this VBScript interactively, run like this: cscript c:\scripts\mac.vbs computername
'  Where "computername" is the netbios name of the computer for which you want the MAC address
' NOTE: Typing "cscript c:\scripts\mac.vbs" with no parameter will run against your local computer
'
' To review results, for example, when done, type notepad c:\scripts\computername-MAC.txt
'  where "computername" is the name of the computer you ran this against.
'
' Assumptions: 
'  1) If running against remote computers, you have rights/privs to run this
'  2) If running on a local workstation, you create or have a "c:\scripts" to save this script and 
'       a "c:\scripts\macinfo" folder to hold the output
'
'  NOTES:
'   I currently run this at login for all users as a useful auditing tool; to do that:
'   1) Simply modify the "strFolder" line to output to a central server shared scripts/output folder 
'      (i.e., strFolder = "\\DomCntrlr1\outshare")
'   2) Then, assuming you have a general login script that runs for every user, modify that script
'      to run this "mac.vbs" script (i.e., cscript \\DomCntrlr1\outshare\mac.vbs)
'   3) Review results by typing "notepad \\DomCntrlr1\outshare\computername-MAC.txt" (where 
'        "computername" is the name of any computer against which this script has run)
'   *) Keep in mind that the output files are APPENDED, meaning this keeps a cumulative record of
'        the Computer Name, MAC, IP and date so, as time goes by, you will see multiple entries
'        in each computer's output file, allowing you to audit when a MAC or IP changed
'
'   To run this against multiple remote computers
'   1) Create a bat script; i.e., c:\script\mac.bat
'   2) Within "mac.bat," call the "mac.vbs" script, with computername as parameter, one line per computer
'      For example:
'       REM - mac.bat
'       cscript c:\scripts\mac.vbs remotepc1
'       cscript c:\scripts\mac.vbs remotepc2
'   *) Review results: notepad c:\scripts\macinfo\remotepc1-MAC.txt (or remotepc2-MAC.txt)
'
'  ALSO: If you want to overwrite the "computername-MAC.txt" file each time, simply use
'  Const ForWriting = 2 (vs. Const ForAppending = 8) & change setobjFile line to "ForWriting" vs. "ForAppending"
'-----------------------------------------------------------------------------------------------

On error resume next
strComputer = Wscript.Arguments.Item(0)

If strComputer = "" Then strComputer="." 'Use "." (local computer) if no computer-name was entered

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

ReDim arrIP(-1)
For Each objItem in colItems
'  Wscript.Echo objItem.MACAddress 'Uncomment if you want to display the info
   macinfo = objItem.MACAddress
   For Each addr In objItem.IPAddress
    ReDim Preserve arrIP(UBound(arrIP)+1)
    arrIP(UBound(arrIP)) = addr
   Next
Next

strIP = Join(arrIP, ",") 

If strComputer = "." Then
  Set objWMIService = GetObject( "winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )

'Grab the netbios name of each computer
  For Each objItem in colItems
     host = objItem.Name
'    WScript.Echo "MAC Address for Computer Name: " & host & " is " & macinfo 'Uncomment to display     
  Next
Else
  host = strComputer
End If

strLine = host & "," & macinfo & "," & strIP & "," & Date

Dim objFile, objFolder, objFSO
Dim strFolder, strFile, strLine
strFolder = "c:\scripts\macinfo"
strFile = "\" & host & "-MAC.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")

set objFile = nothing
' For auditing, we want APPEND mode to track all MAC/IPs for the computer
Const ForAppending = 8

Set objFile = objFSO.OpenTextFile (strFolder & strFile, ForAppending, True)

' Write Computer Name, MAC address, IP Address(es) and Date to a text file
objFile.WriteLine(strLine)
objFile.Close

wscript.quit(0)
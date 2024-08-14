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
Option Explicit
'################################################
' The starting point of execution for this script.
'################################################
Sub Main()
	On Error Resume Next  
	Dim objShell, PATH, bits,Drive,substr
	Dim fso, d, drives,s,count 
	Dim exitcode 
	Dim x, i , j
	exitcode = 0
	bits = 0
	Set objShell = CreateObject("WScript.Shell")		'Create wscript.shell object
	PATH = "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDrives"
	Drive = inputbox("Enter the Drive letter you want to hide")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set drives= fso.Drives
	If UCase(Drive) = "SHOWALL" Then 
		objShell.RegDelete PATH	 'Delete Write registry ,set system default
		If Err = 0 Then 
			Call RefreshExplorer
			Err.clear
		Else 
			wscript.echo  "Operation failed.There is no hidden partition(s) or you do not have administrator permission!"
			Err.clear
		End If 
	Else 
		For i = 1 To Len(Drive)
			substr= Mid(Drive,i,1)
			x = 0
			For j = Asc("A") To Asc("Z")
				If Chr(j) = UCase(substr) Then
					'Varify if  the input letter duplicate
					If InStr(Drive,substr) = InStrRev(Drive,substr)  Then 

							'Varify if the specifid letter exists 
					        If  fso.DriveExists(substr) Then 						
								bits = bits + 2^x 
							Else 
								wscript.echo "The Drive letter '"&UCase(substr)&"' not exsits!Try again!"
								exitcode = 1
								Exit For 
							End If 
					Else 
						wscript.echo "Something dupilcate,please try again"
						exitcode = 1
					End If 
				End If
				x = x + 1
			Next
			If exitcode = 1 Then
				Exit For 
			End If 
		Next
		'Check the exitcode, 
		Dim sNoDrives
		'Read the regsitry 
		sNoDrives = objShell.RegRead(PATH)
		Err.clear
		If exitcode = 0 And drive <> "" Then 
			If bits = 0 Then 
				wscript.echo "Please enter something like 'D,e' "
				Call Main
			Else 
				If sNoDrives = bits Then 
					wscript.echo "The partiton have been hidden."
				Else 
					objShell.RegWrite PATH, bits, "REG_DWORD"		' Write registry,hide the specifid drive
					If Err = 0 Then 
						Call RefreshExplorer
					Else 	
						wscript.echo
						wscript.echo  "Operation failed.Ensure that you have the administrator permission"
						Err.clear 
					End If 
				End If 
			End If 
		ElseIf exitcode = 1 Then 
			Call Main
		End If 
    End If 

End Sub 


Function RefreshExplorer()
'################################################
'Refresh explorer
'################################################
	dim strComputer, objWMIService, colProcess, objProcess 
	strComputer = "."
	'Get WMI object 
	Set objWMIService = GetObject("winmgmts:" _
	  & "{impersonationLevel=impersonate}!\\" _ 
	  & strComputer & "\root\cimv2") 
	Set colProcess = objWMIService.ExecQuery _
	  ("Select * from Win32_Process Where Name = 'explorer.exe'")
	For Each objProcess in colProcess
	   objProcess.Terminate()
	Next 

End Function 

Call Main
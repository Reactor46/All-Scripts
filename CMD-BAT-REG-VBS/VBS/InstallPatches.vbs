Set WshShell = Wscript.CreateObject("Wscript.Shell")

'Root Path e.g. "c:\updates\"
strProgramPath = "<PATH_TO_UPDATE_ROOT" 

strComputer = "."


Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,authenticationLevel=Pkt}!\\" & strComputer & "\root\cimv2")

	Set colSettings = objWMIService.ExecQuery ("SELECT Architecture FROM Win32_Processor")

		For Each objProcessor In colSettings
		    If objProcessor.Architecture = 9 Then 
		    Wscript.Echo "64 bit OS"
			 WshShell.Run strProgramPath & "<SUB_Path_x64>\<File_Name_1>", 1, true
			 WshShell.Run strProgramPath & "<SUB_Path_x64>\<File_Name_2>", 1, true
			 WshShell.Run strProgramPath & "<SUB_Path_x64>\<File_Name_3>", 1, true

		ElseIf objProcessor.Architecture = 0 Then 
		    Wscript.Echo "32 bit OS"
			 WshShell.Run strProgramPath & "<SUB_Path_x86>\<File_Name_1>", 1, true
			 WshShell.Run strProgramPath & "<SUB_Path_x86>\<File_Name_2>", 1, true
			 WshShell.Run strProgramPath & "<SUB_Path_x86>\<File_Name_3>", 1, true
	End If

Next 

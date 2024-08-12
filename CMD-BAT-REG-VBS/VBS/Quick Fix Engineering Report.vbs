' Quick Fix Engineering Report V1.0 2/11/2010

on error resume next

' ********** Get computer name from the user
strComputer=inputbox("Enter Computer Name: ", "QFE")

' ********** Blank the report message
strMsg = ""

' ********** Set computer object 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

if err.number = "-2147217375" then
	' Do nothing
else

	' ********** Check to make sure the computer exists on the network.
	Select Case err.number
		Case 462
			strWarn=MsgBox("Unable to connect to " & strComputer & ".", 48, "QFE")
		Case -2147217394
			strWarn=MsgBox(strComputer & " is not a valid name.", 48, "QFE")
		Case 70
			strWarn=MsgBox(strComputer & " has denied access.", 48, "QFE")
    	Case Else
		

		' ********** Set query for the Computer System and report
		Set colSettings = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
		For Each objComputer in colSettings 
    			select case objComputer.DomainRole
				case 0
	    				strRole = "Standalone Workstation"
				case 1
	    				strRole = "Member Workstation"
				case 2
	    				strRole = "Standalone Server"
				case 3
	    				strRole = "Member Server"
				case 4
	    				strRole = "Backup Domain Controller"
				case 5
	    				strRole = "Primary Domain Controller"
				case else
	    				strRole = "Unknown"
    			end select
    			strMsg = strMsg & "System Name: " & VbTab & VbTab & objComputer.Name & VbCrLf
    			strMsg = strMsg & "System Manufacturer: " & VbTab & objComputer.Manufacturer & VbCrLf
    			strMsg = strMsg & "System Model: " & VbTab & VbTab & objComputer.Model & VbCrLf
    			strMsg = strMsg & "System Role: " & VbTab & VbTab & strRole & VbCrLf
			memSize = round(objComputer.TotalPhysicalMemory / 1073741824,2)
    			strMsg = strMsg & "Total Physical Memory: " & VbTab & memSize & " GB" & VbCrLf
		Next

		' ********** Set query for the Operating System and report
		Set colSettings = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
		For Each objOperatingSystem in colSettings 
    			strMsg = strMsg & "OS Name: " & VbTab & VbTab & objOperatingSystem.Caption & VbCrLf
			strMsg = strMsg & "Version: " & VbTab & VbTab & objOperatingSystem.Version & VbCrLf
			strMsg = strMsg & "Service Pack: " & VbTab & VbTab & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & VbCrLf
			dtmBootup = objOperatingSystem.LastBootUpTime
    			dtmLastBootupTime = CDate(Mid(dtmBootup, 5, 2) & "/" & Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) & " " & Mid (dtmBootup, 9, 2) & ":" & Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, 13, 2))
    			dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
			strMsg = strMsg & "Report Date:" & VbTab & VbTab & now & VbCrLf
			strMsg = strMsg & "Last Boot: " & VbTab & VbTab & dtmLastBootupTime & VbCrLf
			strMsg = strMsg & "System Boot: " & VbTab & VbTab & objoperatingSystem.SystemDrive & VbCrLf
			strMsg = strMsg & "Uptime: " & VbTab & VbTab & dtmSystemUptime & " hours" & VbCrLf
		Next

		' ********** Set the query for the chassis type **********
		Set colChassis = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
		For Each objChassis in colChassis
    			For  Each strChassisType in objChassis.ChassisTypes
        			Select Case strChassisType
          				Case 1
                				strChassisType = "Other"
            				Case 2
                				strChassisType = "Unknown"
            				Case 3
                				strChassisType = "Desktop"
            				Case 4
                				strChassisType = "Low Profile Desktop"
            				Case 5
                				strChassisType = "Pizza Box"
            				Case 6
                				strChassisType = "Mini Tower"
            				Case 7
                				strChassisType = "Tower"
            				Case 8
                				strChassisType = "Portable"
            				Case 9
                				strChassisType = "Laptop"
            				Case 10
                				strChassisType = "Notebook"
            				Case 11
                				strChassisType = "Handheld"
            				Case 12
                				strChassisType = "Docking Station"
            				Case 13
                				strChassisType = "All-in-One"
            				Case 14
                				strChassisType = "Sub-Notebook"
            				Case 15
                				strChassisType = "Space Saving"
            				Case 16
                				strChassisType = "Lunch Box"
            				Case 17
                				strChassisType = "Main System Chassis"
            				Case 18
                				strChassisType = "Expansion Chassis"
            				Case 19
                				strChassisType = "Sub-Chassis"
            				Case 20
                				strChassisType = "Bus Expansion Chassis"
            				Case 21
                				strChassisType = "Peripheral Chassis"
            				Case 22
                				strChassisType = "Storage Chassis"
            				Case 23
                				strChassisType = "Rack Mount Chassis"
            				Case 24
                				strChassisType = "Sealed-Case PC"
            				Case Else
                				strChassisType = "Unknown"
            			End Select
				
				strMsg = strMsg & "Chassis Type: " & VbTab & VbTab & strChassisType & VbCrLf

    			Next
		Next


		' ********** Set the query for the BIOS and report
		Set colSettings = objWMIService.ExecQuery ("Select * from Win32_BIOS")
		For Each objBIOS in colSettings 
    			strMsg = strMsg & "BIOS Version: " & VbTab & VbTab & objBIOS.Version & VbCrLf
    			strMsg = strMsg & "Serial Number: " & VbTab & VbTab & objBIOS.SerialNumber & VbCrLf
		Next

		strMsg = strMsg & VbCrLf

		' ********** Set the query for the Processor(s) and report
		Set colProcessors = objWMIService.ExecQuery ("SELECT * FROM Win32_Processor")
		strMsg = strMsg & "Processors or Cores: " & colProcessors.Count & VbCrLf
		For Each objProcessor in colProcessors
    			strMsg = strMsg & "    Name:         " & objProcessor.Name & VbCrLf
    			strMsg = strMsg & "    Description:  " & objProcessor.Description & VbCrLf
    			strMsg = strMsg & "    Manufacturer: " & objProcessor.Manufacturer & VbCrLf
    			strMsg = strMsg & "    Socket:       " & objProcessor.SocketDesignation & VbCrLf
		Next

		strMsg = strMsg & VbCrLf
		
		
		' ********** Get QFE info **********
		Set colQFE = objWMIService.ExecQuery ("Select * from Win32_QuickFixEngineering")
		strMsg = strMsg & "QFE: " & VbCrLf
		For Each objQFE in colQFE
			strMsg = strMsg & "    Description:  " & objQFE.Description & VbCrLf
			strMsg = strMsg & "    Fix Comments: " & objQFE.FixComments & VbCrLf
			strMsg = strMsg & "    Hotfix ID:    " & objQFE.HotFixID & VbCrLf
			strMsg = strMsg & "    Description:  " & objQFE.Description & VbCrLf
			strMsg = strMsg & "    Installed by: " & objQFE.InstalledBy & VbCrLf
			strMsg = strMsg & "    Installed:    " & objQFE.InstalledOn & VbCrLf
			strMsg = strMsg & "    SP in effect: " & objQFE.ServicePackInEffect & VbCrLf
			strMsg = strMsg & VbCrLf
		next
		
		strMsg = strMsg & VbCrLf
		
		' ********** Get Patch State info **********
		'Set colPatchState = objWMIService.ExecQuery ("Select * from Win32_PatchState")
		'strMsg = strMsg & "Patch State: " & VbCrLf
	'	For Each objPatchState in colPatchState
	'		strMsg = strMsg & "    Product:      " & objPatchState.Description & VbCrLf
	'		strMsg = strMsg & "    Product:      " & objPatchState.Description & VbCrLf
	'	next
		
	'	strMsg = strMsg & VbCrLf

		' ********** Set query for the Logical Disks and report
		Set colSettings = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
		strMsg = strMsg & "Logical Drives: " & VbCrLf
		For Each objLogicalDisk in colSettings 
			
    			strMsg = strMsg & "    Description: " & objLogicalDisk.Name & " " & objLogicalDisk.VolumeName & VbCrLf
			strMsg = strMsg & "         Device Type: " & objLogicalDisk.Description & VbCrLf

		if objLogicalDisk.Description = "3 1/2 Inch Floppy Drive" or objLogicalDisk.Description = "CD-ROM Disc" or objLogicalDisk.Description = "Network Connection" then
			strMsg = strMsg & VbCrLf
		else
			strMsg = strMsg & "         File System: " & objLogicalDisk.FileSystem & VbCrLf
			if objLogicalDisk.MediaType = 12 then
				if objLogicalDisk.VolumeDirty = "True" then
					strMsg = strMsg & "                 *** Run CHKDSK on this volume. ***" & VbCrLf
				end if
			end if

			if objLogicalDisk.Size < 1000000000 then
				partSize = round(objLogicalDisk.Size / 1048576, 2)
				strMsg = strMsg & "         Size:        " & partSize & " MB   " & VbCrLf
			end if
			if objLogicalDisk.Size => 1000000000 then
				partSize = round(objLogicalDisk.Size / 1073741824, 2)
				strMsg = strMsg & "         Size:        " & partSize & " GB   " & VbCrLf
			end if

			if objLogicalDisk.FreeSpace < 1000000000 then
				diskFree = round(objLogicalDisk.FreeSpace / 1048576, 2)
				strMsg = strMsg & "         Free:        " & diskFree & " MB   " & VbCrLf
			end if 
			if objLogicalDisk.FreeSpace => 1000000000 then
				diskFree = round(objLogicalDisk.FreeSpace / 1073741824, 2)
				strMsg = strMsg & "         Free:        " & diskFree & " GB   " & VbCrLf
			end if

			pctFree = int((objLogicalDisk.FreeSpace / objLogicalDisk.Size) * 100)
			strMsg = strMsg & "         % Free:      " & pctFree & VbCrLf

			if pctFree < 15 then
				strMsg = strMsg & "                 *** Free space is less than 15% ***" & VbCrLf
			end if

			strMsg = strMsg & VbCrLf
						
			partSize = 0
			diskFree = 0
			pctFree = 0
		end if

		err.number = 0
					
		next

		' ********** Set query for the Physical Disks and report
		Set colDiskDrive = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
		strMsg = strMsg & "Physical Drives: " & VbCrLf
		For Each objDiskDrive in colDiskDrive 
			if objDiskDrive.Size < 1000000000 then
				diskSize = round(objDiskDrive.Size / 1048576, 2)
				strMsg = strMsg & "    Size:       " & diskSize & " MB   " & VbCrLf
			end if
			if objDiskDrive.Size => 1000000000 then
				diskSize = round(objDiskDrive.Size / 1073741824, 2)
				strMsg = strMsg & "    Size:       " & diskSize & " GB   " & VbCrLf
			end if
    			strMsg = strMsg & "    Caption: " & VbTab & objDiskDrive.Caption & VbCrLf
    			strMsg = strMsg & "    Interface: " & VbTab & objDiskDrive.InterfaceType & VbCrLf
			strMsg = strMsg & "    Partitions: " & VbTab & objDiskDrive.Paritions & VbCrLf
			strMsg = strMsg & "    Status: " & VbTab & objDiskDrive.Status & VbCrLf & VbCrLf
			diskSize = 0
		Next

		' ********** Check for the existence of the "SysInfoCheck" folder then
		' ********** write the file to disk.
		strDirectory = "C:\SysInfoCheck"
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(strDirectory) Then
    			' Procede
		Else
    			Set objFolder = objFSO.CreateFolder(strDirectory)
		End if

		' ********** Calculate date serial for filename **********
		intMonth = month(now)
		if intMonth < 10 then
			strThisMonth = "0" & intMonth
		else
			strThisMonth = intMOnth
		end if
		intDay = Day(now)
		if intDay < 10 then
			strThisDay = "0" & intDay
		else
			strThisDay = intDay
		end if
		strFilenameDateSerial = year(now) & strThisMonth & strThisDay

		Set objFile = objFSO.CreateTextFile(strDirectory & "\" & strComputer & "_" & strFilenameDateSerial & ".txt",True)	
		objFile.Write strMsg & vbCrLf

		' ********** Ask to view file
		strFinish = "Finished collecting information for server: " & strComputer & "." & VbCrLf & VbCrLf & "View file?"
		strAnswer=MsgBox(strFinish, 68, "QFE")
		if strAnswer = 6 then
    			Set objShell = CreateObject("WScript.Shell")
    			objShell.run strDirectory & "\" & strComputer & "_" & strFilenameDateSerial & ".txt"
		end if

	end select

end if
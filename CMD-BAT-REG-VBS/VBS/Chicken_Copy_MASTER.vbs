'THIS SCRIPT CAN ONLY BE USED FOR COPYING FILES.  IT CAN NOT BE USED FOR MIRRORING DATA!
'DO NOT USE IF ROOT DATA IS UPDATED IN ANY WAY AFTER THE FILE HAS BEEN COPIED, FILE CHANGES WILL NOT GET MOVED!
'Another script brought to you by SCRIPT CHICKEN!

On Error Resume Next
'--------------------------------------------Declarations, Objects and Const-------------------------------------------------------------------


Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim strDate1
Dim strTime1
Dim strTime2
Dim myDateString1
Dim myTimeString4
myDateString1 = FormatDateTime(Date(), 1)
myTimeString4 = FormatDateTime(Time(), 4)
strDate = Replace(Date, "/","_")
strTime1 = Left(myTimeString4, 2)
strTime2 = Right(myTimeString4, 2)
Const adVarChar = 200
Const MaxCharacters = 255
Const adFldIsNullable = 32


'---------------------------------------------Set Variables Here--------------------------------------------------------------------------------

SPath = "PATH\TO\ROOT\FOLDER"																															'root folder location
logRoot = "PATH\TO\ROOT\LOG\FOLDER\"																													'root of the log files location
logRetention = 180																																		'rentention period for log files	
configRoot = "PATH\TO\CONFIGFILE\ROOT\"							 																						'root folder where config file is stored (DO NOT PUT FILE NAME IN HERE)																								     		' master_archive file path, txt file with everything ever archived																				' open master archive file to add to it.
rootServer = "\\SERVERorDRIVEPATH"																														'set this to either the current server name or drive path of where the data resides EXAMPLES: "\\servername" "C:\"
destServer = "\\SERVERorDRIVEPATH"																														'set this to where the copied data resides root path EXAMPLES: "\\servername" "D:\"

'--------------------------------------------End Declarations and Sets-------------------------------------------------------------------------- 
 
Set rootFolder = FSO.GetFolder(SPath)                        																							'set object with root folder                                  																							' get size of root folder in bytes

'------------------------------------------Do Copy Job------------------------------------------------------------------------------------------

Set logFile = FSO.OpenTextFile(logRoot & "Chicken__Run_Log_" & strDate & "_" & strTime1 & "-" & strTime2 & ".txt", ForWriting, True)  					'create job log file
logFile.Write "<!------------Chicken Copy Run Log-----by: Script Chicken-------!>" & VbCrLf  

daysBetween()																																			'checks the days between the last run job and the day its ran using the config file
CheckFolder FSO.GetFolder(SPath)                             																							'Check Root of folder for archiving
CheckSubfolders FSO.GetFolder(SPath)					     																							'start parsing through subfolders looking through all subfolders

logFile.Write VbCrLf & VbCrLf & "<!-------------------------Log File End-------------------------!>" & VbCrLf       									'finish writing moved files to the log file
logFile.Close																		  																	'closes the log file
Set config = FSO.CreateTextFile(configRoot & "chicken_config.txt", ForWriting)																			'open the config file to update the date
config.Write date																																		'updates the config file with the current date
config.Close																																			'closes the file
logCleanup()																																			'runs the logfile cleanup
donePopup = objShell.Popup("Files successfully copied.", 10, "Task Success")     							   											'display completion popup


'--------------------------------------------Start Subs and Functions For Archiving-------------------------------------------------------------


		Sub CheckSubFolders(Folder)                          											 												' Find all sub folders within root
			For Each Subfolder in Folder.SubFolders
				CheckFolder(subfolder)													   				 												' for each subfolder, check each  with in object
				CheckSubFolders Subfolder
			Next
		End Sub
		
		Sub CheckFolder(objCurrentFolder)                                       																		'checks files within the specified folder
				FolderModified = FormatDateTime(objCurrentFolder.DateLastModified, "2")
				if DateDiff("d", FolderModified, date) > daysBetween Then
					logFile.Write VbCrLf & "Folder Skipped: " & objCurrentFolder
					else
						For Each objFile In objCurrentFolder.Files                               
						sourceName = objFile
						destName = Replace(sourceName,rootServer,destServer)	
								if FSO.FileExists(destName) then  																						'checks that the file is within the specified time period
								logFile.Write VbCrLf & "File Exists: " & destName
								else
								   mFolder = objCurrentFolder
								   mFolder = Replace(mFolder,rootServer,destServer)                                 									'replaces the root folder path with the destination path
										If Not FSO.FolderExists(mFolder) Then										 									'check to see if root folder is in the destination 
											   objShell.Run "cmd /c mkdir """ & mFolder & """"							 								'if doesn't exist, create folder using command shell
											   WScript.Sleep 200 
											   logFile.Write VbCrLf & "Folder Created! " & mFolder														'wait 2/10 of second before ending If statement
										End If
								   FSO.CopyFile sourceName, destName
								   logFile.Write VbCrLf & "File Copied! " & destName 
								End If
						Next
				End If
		End Sub
		
		
	
		function daysBetween()																															'checks the days between run time and last run time
			If (FSO.FileExists(configRoot & "chicken_config.txt") Then
				Set dateFile = FSO.OpenTextFile(configRoot & "chicken_config.txt", ForReading, True)
				lastRun = dateFile.ReadAll()
				dateFile.Close
				daysBetween = DateDiff("d",lastRun,date)
				Else
					logFile.Write "CANNOT FIND CONFIG FILE!  ENDING SCRIPT!"
					logFile.Close
					Wscript.Quit
			End If
		End Function
		

		
'--------------------------------------------------------Log File Cleanup------------------------------------------------------------------------------

		Sub logCleanup()																																'cleans the log files up
			Dim objFile	
			Set objFolder = FSO.GetFolder(logRoot)
			For Each objFile In objFolder.Files
				FileCreated = FormatDateTime(objFile.DateCreated, "2")
				If DateDiff("d", FileCreated, date) > logRetention Then
				FSO.DeleteFile(objFile)
				End If
			Next
		End Sub

'--------------------------------------------------------End Script------------------------------------------------------------------------------------
		
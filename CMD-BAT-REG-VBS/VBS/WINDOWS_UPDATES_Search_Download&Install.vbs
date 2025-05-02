'Before you attempt to run this sample, note the following:
'•WUA must be installed on the computer. For more information about how to determine the version of WUA that is installed, see Determining the Current Version of WUA..
'•WUA prompts the user to restart the computer if an update requires a restart.
'•The sample can download updates only by using WUA. It cannot download updates from a Software Update Services (SUS) 1.0 server.
 
'Running this sample requires Windows Script Host (WSH). For more information about WSH, see the WSH section of the Platform Software Development Kit (SDK). 
'If the sample is copied to a file named WUA_SearchDownloadInstall.vbs, you can run the sample by opening a Command Prompt window and typing the following command at the command prompt.
 'cscript WUA_SearchDownloadInstall.vbs
 ' or you can create a simple batch file like below and then run that batch file
 'cscript WUA_SearchDownloadInstall.vbs
'pause




Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

WScript.Echo "Searching for updates..." & vbCRLF

Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software'")


WScript.Echo "List of applicable items on the machine:"

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo I + 1 & "> " & update.Title
Next

If searchResult.Updates.Count = 0 Then
	WScript.Echo "There are no applicable updates."
	WScript.Quit
End If

WScript.Echo vbCRLF & "Creating collection of updates to download:"

Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo I + 1 & "> adding: " & update.Title 
    updatesToDownload.Add(update)
Next

WScript.Echo vbCRLF & "Downloading updates..."

Set downloader = updateSession.CreateUpdateDownloader() 
downloader.Updates = updatesToDownload
downloader.Download()

WScript.Echo  vbCRLF & "List of downloaded updates:"

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    If update.IsDownloaded Then
       WScript.Echo I + 1 & "> " & update.Title 
    End If
Next

Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

WScript.Echo  vbCRLF & _
"Creating collection of downloaded updates to install:" 

For I = 0 To searchResult.Updates.Count-1
    set update = searchResult.Updates.Item(I)
    If update.IsDownloaded = true Then
       WScript.Echo I + 1 & "> adding:  " & update.Title 
       updatesToInstall.Add(update)	
    End If
Next

WScript.Echo  vbCRLF & "Would you like to install updates now? (Y/N)"
strInput = WScript.StdIn.Readline
WScript.Echo 

If (strInput = "N" or strInput = "n") Then 
	WScript.Quit
ElseIf (strInput = "Y" or strInput = "y") Then
	WScript.Echo "Installing updates..."
	Set installer = updateSession.CreateUpdateInstaller()
	installer.Updates = updatesToInstall
	Set installationResult = installer.Install()
	
	'Output results of install
	WScript.Echo "Installation Result: " & _
	installationResult.ResultCode 
	WScript.Echo "Reboot Required: " & _ 
	installationResult.RebootRequired & vbCRLF 
	WScript.Echo "Listing of updates installed " & _
	 "and individual installation results:" 
	
	For I = 0 to updatesToInstall.Count - 1
		WScript.Echo I + 1 & "> " & _
		updatesToInstall.Item(i).Title & _
		": " & installationResult.GetUpdateResult(i).ResultCode 		
	Next
End If
		



'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created By : Assaf Miron
' Date : 25/12/2007
' Download MP3Info -  MP3Info Control (c) 1999 WoLoSoft International
' http://www.wolosoft.com
'=*=*=*=*=*=*=*=*=*=*=*=*=
Dim analyzer,fso
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
Const BIF_RETURNONLYFSDIRS = 1 
Const BIF_DONTGOBELOWDOMAIN = 2 
Const BIF_STATUSTEXT = 4 
Const BIF_RETURNFSANCESTORS = 8 
Const BIF_EDITBOX = 16 
Const BIF_VALIDATE = 32 
Const BIF_NEWDIALOGSTYLE = 64 
Const BIF_BROWSEINCLUDEURLS = 128 
Const BIF_BROWSEINCLUDEFILES = &H4000 
Const BIF_SHAREABLE = &H8000 
Const M3UFile = "D:\MyMusic.M3U"
Const ForAppending = 8
AllOptions = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_RETURNFSANCESTORS
AllOptions = AllOptions + BIF_VALIDATE + BIF_NEWDIALOGSTYLE + BIF_BROWSEINCLUDEURLS + BIF_SHAREABLE

Set analyzer = WScript.CreateObject( "MP3Info.Control" )
Set fso = CreateObject("Scripting.FileSystemObject")
analyzer.max_frames = 100
If FSO.FileExists(M3UFile) Then
	Set objFile = FSO.OpenTextFile(M3UFile,ForAppending)
Else
	Set objFile = fso.CreateTextFile(M3UFile)
End If
strPath = BrowseFolder("Select a MP3 Folder:")
ObjFile.WriteLine "#EXTM3U"
AnalyzeFolder(strPath)
Do Until intAnswer <> vbYes
intAnswer = MsgBox("Do you want to add more files?", vbYesNo, "Add more MP3 Files")
	strPath = BrowseFolder( "Select another MP3 folder:")       
'	ObjFile.WriteLine "#EXTM3U"
	AnalyzeFolder(strPath)
Loop
'	

Wscript.echo "Done"
objFile.Close

Function BrowseFolder(strTitle)
Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.BrowseForFolder _
    (WINDOW_HANDLE, strTitle, No_Options)       
If (Not objFolder is nothing) Then       
	Set objFolderItem = objFolder.Self
	BrowseFolder = objFolderItem.Path
Else
	objFile.Close
	Wscript.Quit
End If
End Function

Function AnalyzeFolder( folderspec )
  Dim folder, file, filescollection

  Set folder = fso.GetFolder(folderspec)
  Set CollSubFolders = Folder.SubFolders
  'If CollSubFolders <> Nothing Then
	For Each SubFolder in CollSubFolders
		AnalyzeFolder(SubFolder.Path)
	Next
'  End If
  Set filescollection = folder.Files
  For Each file in filescollection
	If instr(File.Path,".mp3") > 1 Then
		AnalyzeFile( file.Path )
	End If
  Next
End Function

Function AnalyzeFile( filespec ) 
On Error Resume Next
  analyzer.OpenFile( filespec )
  Duration = analyzer.Duration
  Artist = analyzer.Artist
  Title = analyzer.Title
  intDot = Instr(Duration,".")
  intSpace = Instr(Duration," ")
  If intDot > 1 Then
	Duration = Mid(Duration,1,intDot - 1)
  End If
  If intSpace >1 Then
	Duration = Mid(Duration,1,intSpace -1)
  End If
  If (Not Artist = "") OR (Not Title = "") Then
	ObjFile.WriteLine "#EXTINF:" & Duration & "," & Artist & " - " & Title
  Else
	arrFileName = Split(filespec,"\")
	FileName = arrFileName(Ubound(arrFileName))
	FileName = Mid(FileName,1,Len(FileName) - 4)
	ObjFile.WriteLine "#EXTINF:" & Duration & " ," & FileName
  End If
  ObjFile.WriteLine filespec
End Function

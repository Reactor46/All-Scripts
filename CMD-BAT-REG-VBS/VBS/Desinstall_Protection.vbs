Const Key = "HKLM\Software\Protection\"
Const MDP = "HKLM\Software\Protection\MDP"
InputPassword

Sub Debloquer()
    Set WshNetwork = CreateObject("WScript.Network")
    NomMachine = WshNetwork.ComputerName
NomUtilisateur = WshNetwork.UserName
Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists("c:\Protection") Then
   Command1 = "%COMSPEC% /c Echo o| cacls c:\Protection /g " & qq(NomMachine) & ":f administrateurs:f"
   Command2 = "%COMSPEC% /c attrib -s -h -r c:\Protection"
   Result1 = objShell.Run(Command1,0,True)'exécution de la commande sans afficher la console MS-DOS
   Result2 = objShell.Run(Command2,0,True)'exécution de la commande sans afficher la console MS-DOS
   If Result <> 0 Then
     MsgBox "Permissions sur le dossier non fait",16,"Permissions sur le dossier non fait"
   End If
    End if 
End Sub
'==============
Function qq(strIn)
    qq = Chr(34) & strIn & Chr(34)
End Function
'==============
Function Scramble (strText, lngSeed)
     Dim L,intRand,bytASC
     '---- Force seeded random mode 
     Rnd(-1)
     '---- Set (positive) seed 
     Randomize ABS(lngSeed)
     '---- Scan through string
     For L = 1 To Len(strText)
         '---- Get ASC of char
         bytASC=Asc(Mid(strText, L))
         '---- Fix for quotes (tilde to quote)
         If bytASC=126 then bytASC=34
         '---- Add a random value from -80 to 80, encode\decode is decided by the seed's sign
         intRand = bytASC + ((Int(Rnd(1) * 160) - 80) * SGN(lngSeed)) 
         '---- Cycle char between 32 and 125 (with carry)
         If intRand <= 31 Then 
             intRand = 125 - (31 - intRand)
         ElseIf intRand >= 126 Then
             intRand = 32 + (intRand - 126)
         End If
         '---- Fix for quotes (quote to tilde)
         If intRand=34 then intRand=126
          '---- Output string
         Scramble = Scramble & Chr(intRand)
     Next
End Function
'========================
Sub InputPassword()
Const ForWriting = 2
Const ForAppending = 8
Dim Ws,Password,itemtype,LireMDP
If  Not RegExists(MDP) Then MsgBox "Le Mot de passe n'est pas installé sur ce système !",48,"Mot de passe non installé !" :Wscript.Quit(0): End If
Titre=" Protection Dossier © Hackoo © 2012 "
Set Ws = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
itemtype = "REG_SZ"
Set colItems = GetObject("winmgmts:root\cimv2").ExecQuery("Select ScreenHeight, ScreenWidth from Win32_DesktopMonitor Where ScreenHeight Is Not Null And ScreenWidth Is Not Null") 
     
    For Each objItem in colItems 
        intHorizontal = objItem.ScreenWidth
        intVertical = objItem.ScreenHeight
    Next 
On error resume next  
    Dim objExplorer : Set objExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
    With objExplorer
        .Navigate "about:blank"  
        .ToolBar = 0
        .Left = (intHorizontal-300) / 2
        .Top = (intVertical-300) / 2
        .StatusBar = 0
        .Width = 320
        .Height = 190
        .Visible = 1   
        .Resizable = 0	
		.MenuBar = 0
        .Document.Title = "Mot de Passe © Hackoo ************************"
        Dim strHTML : strHTML = "<center><h3 style='color:Red'>Entrez Votre Mot de Passe</h3>"
		strHTML = strHTML &"<body bgcolor='#FFFFD2' scroll='no'>"
        strHTML = strHTML & "<input type='password' name='txt_Password' size='30'><br>"
        strHTML = strHTML & "<br><button type='submit' style='font-family:Verdana;font-size:14px;height:30px;Width:180px;' id='btn_Exit' onclick=" & Chr(34)& "VBScript:me.Value='AUTENTIFICATION...'" & Chr(34)& " title='Vérifier le mot de passe...'>Envoyer</button></body></center>"
       .Document.Body.InnerHTML = strHTML
	   .Document.Body.Style.overflow = "auto"
	   .Document.body.style.backgroundcolor="lightblue"
	   .Document.All("txt_Password").Focus
    End With

    Do While (objExplorer.Document.All.btn_Exit.Value = "Envoyer")
        Wscript.Sleep 250
        If objExplorer.Document.All.btn_Exit.Value = "1" Then  
            objExplorer.Quit
            Set objExplorer = Nothing
            Exit Sub
        End If
    Loop
    Password = Trim(objExplorer.document.GetElementByID("txt_Password").Value)
objExplorer.Quit
Set objExplorer = Nothing

  If Scramble(Password,2011) = WS.RegRead(MDP) Then 'LireMDP then
     Question = MsgBox ("Voulez-vous accéder à votre Dossier protégé ?",VBYesNO+VbQuestion,Titre)
   If Question = VbYes then
   Call Debloquer()
   WS.RegDelete Key
   Call NO_Disque_Virtuel()
   Call NO_NameSpace()
   Call NO_CheckProtection()
   MsgBox "Le Programme est désormais désinstallé de votre système !",64,Titre
         End If
      Else
      Call Bloquer
         Set Voix = CreateObject("SAPI.Spvoice")	   
         Voix.Speak "PASSWORD INCORRECT AND PERMISSION DENIED TO ACCESS TO THIS FOLDER."
     Msgbox "! كلمة مرور غير صحيحة للوصول إلى هذا المجلد" & vbCr & vbCr &_
     "MOT DE PASSE INCORRECT ET PERMISSION REFUSEE D'ACCEDER A CE DOSSIER" & vbCr & VbCr & _
    "PASSWORD INCORRECT AND PERMISSION DENIED TO ACCESS TO THIS FOLDER",16,"MOT DE PASSE INCORRECT Hackoo © 2012 !"
      End If
End Sub
'--------------------Fin du InputPassword-------------
Function Explorer(File)
    Set ws=CreateObject("wscript.shell")
    ws.run "Explorer " & File '& "\"
end Function
'==================
Function RegExists(value)
   On Error Resume Next ' Sans cette instruction, une erreur se produit si MDP n'existe pas(val="")
   Set WS = CreateObject("WScript.Shell")
   val = WS.RegRead(value)
   RegExists = (Err.number <> -2147024893) And (Err.number <> -2147024894) And val<> ""
End Function

sub Bloquer()
Set Ws = CreateObject("WScript.Shell")
'Set WshNetwork = CreateObject("WScript.Network")
Set ProcessEnv = Ws.Environment("Process")
 NomMachine = ProcessEnv("COMPUTERNAME") 'WshNetwork.ComputerName
 NomUtilisateur = ProcessEnv("USERNAME") 'WshNetwork.UserName
 'MsgBox NomMachine,64,"Nom de la Machine"
 'MsgBox NomUtilisateur,64,"Nom de l'utilisateur"
Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists("c:\Protection") Then
Command1 = "%COMSPEC% /c attrib +s +h +r c:\Protection"
Command2 = "%COMSPEC% /c Echo o| cacls c:\Protection /p " & qq(NomUtilisateur) & ":n administrateurs:n"
 Result1 = objShell.Run(Command1,0,True)'exécution de la commande sans afficher la console MS-DOS
 Result2 = objShell.Run(Command2,0,True)'exécution de la commande sans afficher la console MS-DOS
If Result2 <> 0 Then
   MsgBox "Permissions sur le dossier non fait",16,"Permissions sur le dossier non fait"
End If
end if
End Sub

Sub NO_Disque_Virtuel()
Set objShell = CreateObject("Wscript.Shell")
Command = "%COMSPEC% /C SUBST Z: /D"
Result = objShell.Run(Command,0,True)'exécution de la commande sans afficher la console MS-DOS
End Sub

Sub NO_NameSpace()
On Error Resume Next
bureau="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
mycomputer="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
ROOT="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
Set Ws = CreateObject("Wscript.Shell")
Ws.Regdelete bureau
Ws.Regdelete mycomputer
'Ws.Regdelete ROOT
KillExplorer = Ws.Run("cmd /C taskkill /f /im explorer.exe",0,TRUE)
StartExplorer = Ws.Run("cmd /C start explorer.exe",0,TRUE)
Explorer "c:\Protection\"
end Sub

Sub NO_CheckProtection()
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("WScript.Shell")
startupPath = Shell.SpecialFolders("AllUsersStartup")
If FSO.FileExists(startupPath & "\CheckProtection.vbs") Then
FSO.DeleteFile startupPath & "\CheckProtection.vbs",True
end if
End Sub
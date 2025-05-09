'VBScriptكيفية إنشاء مجلد وحمايته بواسطة كلمة مرور مع ل 
'© Hackoo ديسمبر 2011
'هذا البرنامج بإنشاء مجلد باسم "حماية" ، ثم قال انه يعطي الإذن
'رفض الوصول ، لذلك لا يمكنك إعادة تسمية أو فتح أو الكتابة أو القراءة أو حذف هذا المجلد
'اختبار على ويندوز 7 64 بت النسخة الفرنسية
'--------------------------------------------Description en Français-------------------------------------------
'Comment créer un dossier et le protéger par mot de passe par vbscript © Hackoo Décembre 2012
'ce vbscript crée un dossier nommé "Protection" puis il lui accorde une permission
'd'un accées refusé, donc vous ne pouvez pas ni renommer,ni ouvrir, ni écrire,ni lire,ni supprimer ce dossier 
'Testé sous Windows 7 64-bits Version Française
'Mise à jour le 22/03/2012
'--------------------------------------------Description in English --------------------------------------------
'How to create a folder and protect it by password with vbscript © Hackoo December 2012
'This vbscript creates a folder named "Protection" then he gives permission
'access denied, so you can not rename or open, or write, or read or delete this folder
'Tested on Windows 7 64-bit French version
'updated on 22/03/2012
'---------------------------------------------------------------------------------------------------------------
'-------------------------------------------Programme Principal--------------------------------------------
 Titre=" Protection Dossier © Hackoo © 2012 "
 MDP = "HKLM\Software\Protection\MDP"
 If RegExists(MDP) Then
 Call InputPassword
 Else
 Call Setup_Password()
end if

sub Bloquer()
Set objShell = CreateObject("Wscript.Shell")
Set ProcessEnv = objShell.Environment("Process")
 NomMachine = ProcessEnv("COMPUTERNAME") 
 NomUtilisateur = ProcessEnv("USERNAME") 
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
 
sub Debloquer()
Set WshNetwork = CreateObject("WScript.Network")
 NomMachine = WshNetwork.ComputerName
 NomUtilisateur = WshNetwork.UserName
Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists("c:\Protection") Then
Command1 = "%COMSPEC% /c Echo o| cacls c:\Protection /g " & qq(NomUtilisateur) & ":f administrateurs:f"
'Command2 = "%COMSPEC% /c attrib -s -h -r c:\Protection"
 Result1 = objShell.Run(Command1,0,True)'exécution de la commande sans afficher la console MS-DOS
 'Result2 = objShell.Run(Command2,0,True)'exécution de la commande sans afficher la console MS-DOS
If Result <> 0 Then
   MsgBox "Permissions sur le dossier non fait",16,"Permissions sur le dossier non fait"
End If
End if 
End Sub

Function qq(strIn)
    qq = Chr(34) & strIn & Chr(34)
End Function

'---------------------------------Fonction Scramble--------------------------------------
'Thanks to the Author of this Function © AMBience
'C'est une Fonction de Cryptage trouvé dans ce lien:
'http://www.visualbasicscript.com/Tiny-text-encryption-m83948.aspx
' strText = String to encrypt\decrypt
' lngSeed = Long number for the random seed (key)
' Returns a string
' To Encrypt:- Send the plain text with a positive seed number (1-2147483647)
' To Decrypt:- Send the encrypted text with the same number but negative
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
'-----------------------------------Fin de la Fonction Scramble--------------------------------------
Function RegExists(value)
 On Error Resume Next
 Set WS = CreateObject("WScript.Shell")
 val = WS.RegRead(value)
 If (Err.number = -2147024893) or (Err.number = -2147024894) Then
 RegExists = False
 Else
 RegExists = True
 End If
 End Function
 '----------------------------------------------------------------------------------------------------
Sub Setup_Password()
Dim Ws,Password,MDP,itemtype,LireMDP
Set Ws = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Voix = CreateObject("SAPI.Spvoice")
If Not FSO.FolderExists("c:\Protection") Then
FSO.CreateFolder ("c:\Protection")
end if
MDP = "HKLM\Software\Protection\MDP"
itemtype = "REG_SZ"
VIDE=True
While VIDE
If Password="" Then
  Set colItems = GetObject("winmgmts:root\cimv2").ExecQuery("Select ScreenHeight, ScreenWidth from Win32_DesktopMonitor Where ScreenHeight Is Not Null And ScreenWidth Is Not Null") 
     For Each objItem in colItems 
        intHorizontal = objItem.ScreenWidth
        intVertical = objItem.ScreenHeight
    Next 
   On error resume next
    Dim objExplorer : Set objExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
    With objExplorer
        .Offline =True
        .Navigate "about:blank"  
        .ToolBar = 0
        .StatusBar = 0
        .Width = 370
        .Height = 280
        .Visible = 1   
        .Resizable = 0	
		.MenuBar = 0
        .Document.Title = "Setup du Mot de Passe © Hackoo ******"
        Dim strHTML : strHTML = "<center><h3 style='color:Red'>Choisissez Votre Mot de Passe</h3>"
		strHTML = strHTML &"<body bgcolor='#FFFFD2' scroll='no'>"
        strHTML = strHTML & "<input type='password' name='txt_Password1' size='30'>"
		strHTML = strHTML & "<h3 style='color:Red'>Retapez Votre Mot de Passe</h3>"
		strHTML = strHTML & "<input type='password' name='txt_Password2' size='30'><br>"
        strHTML = strHTML & "<br><button type='submit' style='font-family:Verdana;font-size:14px;height:30px;Width:180px;' id='btn_Exit' onclick=" & Chr(34)& "VBScript:me.Value='Enregistrement....'" & Chr(34)& " title='Enregistrement....'>Envoyer</button></body></center>"
       .Document.Body.InnerHTML = strHTML
	   .Document.Body.Style.overflow = "auto"
	   .Document.body.style.backgroundcolor="lightblue"
	   .Document.All("txt_Password1").Focus
	   End With
    Do While (objExplorer.Document.All.btn_Exit.Value = "Envoyer")
        Wscript.Sleep 250
        If objExplorer.Document.All.btn_Exit.Value = "1" Then  
            objExplorer.Quit
            Set objExplorer = Nothing
            Exit Sub
        End If
    Loop
	Password1=objExplorer.document.GetElementByID("txt_Password1").Value
	Password2=objExplorer.document.GetElementByID("txt_Password2").Value
       If Password1 = Password2 and Password1 = "" Then
        MsgBox "اختيار كلمة المرور فارغة" & vbcr &_
               "الرجاء اختيار كلمة مرور غير فارغة شكرا"& vbcr & vbcr &_
               "Le mot passe choisi est vide !"& vbcr &_
               "Veuillez SVP Choisir un Mot de Passe non vide "& vbcr &_
               "Merci !"& vbcr & vbcr &_
               "The password chosen is empty !"& vbcr &_
               "Please choose a password is not empty "& vbcr &_
                "Thanks !",48,"Mot de Passe vide © Hackoo ******"
       end if
If Password1 = Password2 and Password1 <> "" Then
    Password = objExplorer.document.GetElementByID("txt_Password2").Value
	PasswordCrypt = Scramble(Password,2011)
	Call Disque_Virtuel()
	Call CopyMyscript()
	Call SetupNameSpace()
	MsgBox PasswordCrypt & " : كلمة المرور المشفرة هي " & vbcr &_
	"Votre Mot de Passe Crypté est: " & PasswordCrypt & vbcr&_
    "Encrypted Password is: "& PasswordCrypt ,64,"Mot de Passe Crypté"
	Msgbox "كلمة السر الخاصة بك في كلمة ""{"& Password &"}"" حفظه في مكان جيد! هذا هو السبيل الوحيد لفتح حماية ملف"& vbcr & vbcr &_
	"VOTRE MOT DE PASSE EN CLAIR EST  ""{"&Password&"}""  SAUVEGARDER LE BIEN ! C'EST LE SEUL MOYEN POUR DEBLOQUER   LE DOSSIER PROTECTION !" &vbcr&vbcr&_
	"YOUR PASSWORD IS IN CLEAR ""{"& Password &"}"" SAVE IT IN A GOOD PLACE ! THIS IS THE ONLY WAY TO UNLOCK THE FILE PROTECTION !",64,"MOT DE PASSE INSTALLE Hackoo © 2012 !"
	else
	Voix.Speak "The two passwords do not match !"
	MsgBox " ! كلمتي سر ليست متطابقة"& vbcr & vbcr &_
	"Les deux mots de passe ne sont pas identiques !" & vbcr & vbcr &_
	"The two passwords do not match !",16,"Mot de Passe Erroné © Hackoo ******"
	end if
    If Password <>"" Then 
      VIDE=False
       Ws.RegWrite MDP, PasswordCrypt, itemtype
	   Call Bloquer()
 Voix.Speak " ! حماية المجلد أنشاء ومحمي بنجاح! وبالطبع يمكنك نسخ المجلدات والملفات في ذلك لحمايتهم"& vbcr & vbcr &_
 "Folder Protection Created and Protected Sucessfully and of course you can copy your folders and files in it to protect them !"
 MsgBox " ! حماية المجلد أنشاء ومحمي بنجاح! وبالطبع يمكنك نسخ المجلدات والملفات في ذلك لحمايتهم"& vbcr & vbcr &_
 "Le Dossier Protection est désormais créé et protégé avec succès ! et vous pouvez copier vos dossiers et vos fichiers dans ce dernier pour les protéger !"& vbcr & vbcr &_
 "Folder Protection Created and Protected Sucessfully ! and of course you can copy your folders and files in it to protect them !",64,Titre
    End if 
End if
Call WriteCheckProtection()
    objExplorer.Quit
	Set objExplorer = Nothing
Wend
KillExplorer = Ws.Run("cmd /C taskkill /f /im explorer.exe",0,TRUE)
Call InputPassword()
StartExplorer = Ws.Run("cmd /C start explorer.exe",0,TRUE)
end Sub

'--------------------------------InputPassword-------------------------
Sub InputPassword()
Const ForWriting = 2
Const ForAppending = 8
Dim Ws,Password,MDP,itemtype,LireMDP
Titre=" Protection Dossier © Hackoo © 2012 "
Set Ws = CreateObject("Wscript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
MDP = "HKLM\Software\Protection\MDP"
itemtype = "REG_SZ"
Set colItems = GetObject("winmgmts:root\cimv2").ExecQuery("Select ScreenHeight, ScreenWidth from Win32_DesktopMonitor Where ScreenHeight Is Not Null And ScreenWidth Is Not Null") 
     
    For Each objItem in colItems 
        intHorizontal = objItem.ScreenWidth
        intVertical = objItem.ScreenHeight
    Next 
 On error resume next  
    Dim objExplorer : Set objExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
    With objExplorer
        .Offline = True
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
        .Document.Title = "Mot de Passe © Hackoo *************************** "
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
    Password = objExplorer.document.GetElementByID("txt_Password").Value
	PassowrdCrypt = Scramble(Password,2011)
	If objExplorer.Document.All.btn_Exit.Value = "AUTENTIFICATION..." Then 
	objExplorer.Quit
	Set objExplorer = Nothing
	End if

If  RegExists(MDP) Then
	 LireMDP = ws.RegRead(MDP)
	 LireMDP = Scramble(LireMDP,-2011)
If Password = LireMDP then
Question = MsgBox ("هل تريد الوصول إلى المجلد الخاص بالحماية؟" & vbcr &_
"Voulez-vous accéder à votre Dossier protégé ? " &vbcr &_
"Do you want to access your protected folder ?",VBYesNO+VbQuestion,Titre)
 If Question = VbYes then
 Call Debloquer()
 Call Disque_Virtuel()
	 Explorer("Z:")
	 else
 Call Bloquer
 Call NO_Disque_Virtuel()
 end if
else
   Call Bloquer
   Call NO_Disque_Virtuel() 
    Set Voix = CreateObject("SAPI.Spvoice")	   
    Voix.Speak "PASSWORD INCORRECT AND PERMISSION DENIED TO ACCESS TO THIS FOLDER."
	Msgbox "! كلمة مرور غير صحيحة للوصول إلى هذا المجلد" & vbCr & vbCr &_
    "MOT DE PASSE INCORRECT ET PERMISSION REFUSEE D'ACCEDER A CE DOSSIER !" & vbCr & vbCr &_
	"PASSWORD INCORRECT AND PERMISSION DENIED TO ACCESS TO THIS FOLDER !",16,"MOT DE PASSE INCORRECT Hackoo © 2012 !"
end if
end if
end sub
'--------------------Fin du InputPassword-------------

Function Explorer(Dir)
    Set ws=CreateObject("wscript.shell")
    ws.run "Explorer.exe "& Dir & "\"
end Function

sub WriteCheckProtection()
dim shell,startupPath,link,temp,FSO
Set Shell = CreateObject("WScript.Shell")
startupPath = Shell.SpecialFolders("AllUsersStartup")
Set FSO = CreateObject("Scripting.FileSystemObject")
set f = FSO.OpenTextFile(startupPath & "\CheckProtection.vbs",2,True)
f.writeline "sub Bloquer()"
f.writeline "Set Ws = CreateObject(""WScript.Shell"")"
f.writeline "Set ProcessEnv = Ws.Environment(""Process"")"
f.writeline "NomMachine = ProcessEnv(""COMPUTERNAME"")" 
f.writeline "NomUtilisateur = ProcessEnv(""USERNAME"")"
f.writeline "Set objShell = CreateObject(""Wscript.Shell"")"
f.writeline "Set objFSO = CreateObject(""Scripting.FileSystemObject"")"
f.writeline "If objFSO.FolderExists(""c:\Protection"") Then"
f.writeline "Command1 = ""%COMSPEC% /c attrib +s +h +r c:\Protection"""
f.writeline "Command2 = ""%COMSPEC% /c Echo o| cacls c:\Protection /p "" & qq(NomUtilisateur) & "":n administrateurs:n"""
f.writeline "Result1 = objShell.Run(Command1,0,True)"
f.writeline "Result2 = objShell.Run(Command2,0,True)"
f.writeline "end if"
f.writeline "End Sub"
f.writeline "Function qq(strIn)"
f.writeline "qq = Chr(34) & strIn & Chr(34)"
f.writeline "End Function"
f.writeline "Call Bloquer()"
End Sub

Sub Disque_Virtuel()
Set objShell = CreateObject("Wscript.Shell")
Command = "%COMSPEC% /C SUBST Z: C:\Protection"
Result = objShell.Run(Command,0,True)'exécution de la commande sans afficher la console MS-DOS
End Sub

Sub NO_Disque_Virtuel()
Set objShell = CreateObject("Wscript.Shell")
Command = "%COMSPEC% /C SUBST Z: /D"
Result = objShell.Run(Command,0,True)'exécution de la commande sans afficher la console MS-DOS
End Sub

Sub CopyMyscript()
Set fso = CreateObject("Scripting.FileSystemObject")
Set Ws = CreateObject("Wscript.shell")
AppData = Ws.ExpandEnvironmentStrings("%AppData%")
Monscript = WScript.ScriptFullName
cible = AppData &"\"
if (not fso.fileexists(cible & Monscript)) then
		fso.copyfile Monscript ,cible, True
		end if
End sub

Sub SetupNameSpace()
Dim ROOT,Namespace,cmd,InProcServer32,shellex,shellfolder,bureau,mycomputer,WSH,DefaultIcon,AppData
ROOT="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
DefaultIcon="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\DefaultIcon\"
cmd="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\Shell\Mot de Passe\Command\"
InProcServer32="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\InProcServer32\"
shellex="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\ShellEx\PropertySheetHandlers\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
shellfolder="HKEY_CLASSES_ROOT\CLSID\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\ShellFolder\"
bureau="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"
mycomputer="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{FD4DF9E0-E3DE-11CE-BFCF-ABCD1DE12345}\"

Set WSH = CreateObject("Wscript.shell")
AppData = WSH.ExpandEnvironmentStrings("%AppData%")
WSH.regwrite ROOT ,"Dossier Système Protégé © Hackoo"
WSH.regwrite DefaultIcon ,"%windir%\system32\shell32.dll,47"
WSH.regwrite shellex ,""
WSH.regwrite cmd  ,"wscript.exe "& qq(AppData) &"\"& wscript.scriptName
WSH.regwrite InProcServer32 ,"shell32.dll"
WSH.regwrite InProcServer32 &"\ThreadingModel","Apartment","REG_SZ"
WSH.regwrite shellfolder ,""
WSH.regwrite shellfolder &"\Attributes" ,"0","REG_DWORD"
WSH.regwrite bureau,""
WSH.regwrite mycomputer,""
End Sub
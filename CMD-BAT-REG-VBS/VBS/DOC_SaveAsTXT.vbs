'######################################################################### 
'  Script name:       DOC_SaveAsTXT.vbs 
'  Created on:        05.10.2011
'  Author:            Dennis Hemken 
'  Purpose:           Opens an existing Microsoft Word Document 
'                     and then saves the file in TXT format. 
'######################################################################### 
 
Dim AppWord  
Dim OpenDocument 
Const docTXT = 2 
 
Set AppWord = CreateObject("Word.Application") 
 
AppWord.Visible = True 
 
Set OpenDocument = AppWord.Documents.Open("C:\Concepts\Temp.doc") 
     
OpenDocument.SaveAs "C:\Concepts\TXT\Description", docTXT 
 
OpenDocument.Close 
Set OpenDocument = Nothing 
 
AppWord.Quit 
Set AppWord = Nothing
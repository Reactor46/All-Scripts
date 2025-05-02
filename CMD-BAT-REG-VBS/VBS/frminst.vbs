Install
Sub Install
'Get the folder
sFolder = InputBox("Please enter the URL to the folder: example http://serverName/public/foldername","Setup Instructions")
If Trim(sFolder) = "" Then
    Exit Sub
End If

'Create the app folder, pointing the SCR to the Resources subfolder
Set oDest = CreateObject("CDO.Folder")
oDest.Fields("urn:schemas-microsoft-com:exch-data:schema-collection-ref") = sFolder + "\Resources"
oDest.Fields("DAV:contentclass") = "urn:content-classes:folder"
oDest.Fields.Update
oDest.DataSource.SaveTo sFolder
'Create the Resources folder and make it invisible
Set oDest = CreateObject("CDO.Folder")
oDest.Fields("DAV:ishidden") = True
oDest.Fields.Update
oDest.DataSource.SaveTo sFolder +  "/Resources"

'Fill the Resources folder with form registrations

Set oCon = CreateObject("ADODB.Connection")
oCon.ConnectionString = sFolder + "/resources"
oCon.Provider = "ExOledb.Datasource"
oCon.Open

'----------------------------------------------------
'Register the default page for the folder
Set oRec = CreateObject("ADODB.Record")
oRec.Open "default.reg", oCon, 3, 0
oRec.Fields("DAV:contentclass") = "urn:schemas-microsoft-com:office:forms#registration"
oRec.Fields("urn:schemas-microsoft-com:office:forms#contentclass") = "urn:content-classes:folder"
oRec.Fields("urn:schemas-microsoft-com:office:forms#cmd") = "*"
oRec.Fields("urn:schemas-microsoft-com:office:forms#formurl") = "phoneup.asp"
oRec.Fields("urn:schemas-microsoft-com:office:forms#executeurl") = "phoneup.asp"
oRec.Fields.Update
oRec.Close

' Upload Page
Const adDefaultStream = -1
Set stm = CreateObject("ADODB.Stream")
set Rec = CreateObject("ADODB.Record")
Rec.Open sFolder + "\Resources\phoneup.asp",oCon ,3,0
Set Stm = Rec.Fields(adDefaultStream).Value
stm.charset = "us-ascii"
stm.loadfromfile "d:\phoneup.asp"
stm.flush
msgbox("done")

End Sub


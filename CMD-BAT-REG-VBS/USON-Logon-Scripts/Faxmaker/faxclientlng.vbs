'*****************************************************************
'Translation script by Stefan Engelbert (GFI Software Ltd.)
'Syntax: cscript translate.vbs <MSIfile> <XMLfile>
'Files should be located in the same directory
'*****************************************************************
On Error Resume Next


If Wscript.Arguments.Count=2 then
	if ((checkfile(folder&Wscript.Arguments(0)&".msi")=1)  and (checkfile(folder&Wscript.Arguments(1)&".xml")=1)) Then
		Const msiOpenDatabaseModeTransact = 1
		Dim openMode : openMode = msiOpenDatabaseModeTransact
		Dim databasePath:databasePath = folder&Wscript.Arguments(0)&".msi"
		Dim installer : Set installer = Nothing:CheckError
		Dim xmlDoc
		Dim currNode
		Dim currNode1
		Dim currNode2
		Dim query
		Dim record
		Dim msitablename
		Dim tables,tablename,subtablename,wert,font
		Set installer = Wscript.CreateObject("WindowsInstaller.Installer") :CheckError
		Set Record = Installer.CreateRecord(1):CheckError
		
		Dim record2
		Dim record2Size		
		record2Size = 0     ' 0 means it was not created yet
		
		Dim database : Set database = installer.OpenDatabase(databasePath, openMode) :CheckError
		Set xmlDoc = CreateObject("Msxml.DOMDocument")
		
		xmlDoc.async = false
		xmlDoc.load(folder&Wscript.Arguments(1)&".xml")
		Set objNodeList = xmlDoc.getElementsByTagName("table")
		tables= objNodeList.length
	
		For i = 0 To (tables - 1)
			
		  	tablename=objNodeList.item(i).getAttribute("id")
			
  			Set currNode = objNodeList.Item(i).getElementsByTagName("*")
  			For m = 0 To (currNode.length - 1)
  			
  			    dim actualNode 
				set actualNode = currNode.item(m)
				
				msitablename=tablename				    			    
			    
			    if actualNode.nodeName="record" then
			        ' determine number of attributes in record			        
			        fieldCount = actualNode.childNodes.length
			        
			        if (fieldCount + 1) > record2Size then
			            record2Size = fieldCount + 1			            
			            set record2 = Installer.CreateRecord(record2Size)
			        end if
			        				        
			        ' build query
			        Dim baseQuery
			        Dim pkName, pkValue

		            baseQuery = "UPDATE `"&msitablename&"` SET "
		            set pkName = nothing			            
			        				        
			        For idx=0 to fieldCount-1
			            Dim fieldInfo
			            set fieldInfo = actualNode.childNodes.item(idx)
			            Dim colname, colVal
			            colname = fieldInfo.getAttribute("id")
			            record2.stringdata(idx+1) = fieldInfo.text			            
			            
			            if idx = 0 then
			                pkName = colname
		                    pkValue = fieldInfo.text
		                end if
			            
			            baseQuery = baseQuery&"`"&msitablename&"`.`"&colname&"`=?"
			            if idx <> fieldCount-1 then
			                baseQuery = baseQuery&", "
			            end if
			        next
			        if pkName is not nothing then
			            'record2.stringdata(actualNode.Attributes.length) = pkValue
			            baseQuery = baseQuery&" WHERE `"&msitablename&"`.`"&pkName&"`='"&pkValue&"'"
			        end if				       
			        
			        'MsgBox baseQuery&chr(10)&" '"&record2.stringdata(0)&"' '"&record2.stringdata(1)&"' '"&record2.stringdata(2)&"' '"&record2.stringdata(3)&"'"
				    Set view = database.OpenView(baseQuery)
				    'view.Execute : CheckError
				    View.Execute record2
				end if
				
				' the following code handles the modules and other tables
  				If currNode.Item(m).getElementsByTagName("*").length>0 Then
  					Set subNode = currNode.Item(m).getElementsByTagName("*")
  					feld=currNode.item(m).getAttribute("id")
				Else
				    subtablename=currNode.item(m).getAttribute("id")
				    font=currNode.item(m).getAttribute("FONT")
				    wert=currNode.item(m).text						
					
				    if ((wert<>"") and (tablename<>"") and(tablename<>subtablename)) then
					    if instr(font,"Arial14")>0 then 
						    font=replace(font,"Arial14","{&Arial14}")
					    end if
					    if instr(font,"TimesItalicBlue10")>0 then 
						    font=replace(font,"TimesItalicBlue10","{\TimesItalicBlue10}")
					    end if
					    if instr(font,"[WiseCRLF]")>0 then
						    font=replace(font,"[WiseCRLF]",chr(17)&chr(25))
					    end if
                        
                        If msitablename="Shortcutname" Then 
                    	    msitablename="Shortcut"
                        End If
                        If msitablename="Shortcutdescription" Then 
                    	    msitablename="Shortcut"
                        End If
						
					    if feld<>"" Then
						    Record.StringData(1)=CStr(font&wert)
						    'query = "UPDATE "&tablename&" SET "&tablename&".Text='"&font&wert&"' WHERE "&tablename&".Dialog_='"&feld&"' AND "&tablename&".Control='"&subtablename&"'"
						    query = "UPDATE "&tablename&" SET "&tablename&".Text=? WHERE "&tablename&".Dialog_='"&feld&"' AND "&tablename&".Control='"&subtablename&"'"
					    Else
						    If msitablename="Shortcut" Then
							    If tablename="Shortcutname" Then
								    Record.StringData(1)=CStr(font&wert)
								    'query = "UPDATE "&msitablename&" SET "&msitablename&".Name='"&font&wert&"' WHERE "&msitablename&".Shortcut='"&subtablename&"'"
								    query = "UPDATE "&msitablename&" SET "&msitablename&".Name=? WHERE "&msitablename&".Shortcut='"&subtablename&"'"
							    End If
							    If tablename="Shortcutdescription" Then
								    Record.StringData(1)=CStr(font&wert)
								    'query = "UPDATE "&msitablename&" SET "&msitablename&".Description='"&font&wert&"' WHERE "&msitablename&".Shortcut='"&subtablename&"'"
								    query = "UPDATE "&msitablename&" SET "&msitablename&".Description=? WHERE "&msitablename&".Shortcut='"&subtablename&"'"
							    End If
						    Else
							    If tablename="ActionText" Then
								    Record.StringData(1)=CStr(font&wert)
								    'query = "UPDATE "&msitablename&" SET "&msitablename&".Value='"&font&wert&"' WHERE "&msitablename&".Property='"&subtablename&"'"
								    query = "UPDATE "&msitablename&" SET "&msitablename&".Description=? WHERE "&msitablename&".Action='"&subtablename&"'"	
								End if
							    If tablename="Error" Then
									Record.StringData(1)=CStr(font&wert)
									'query = "UPDATE "&msitablename&" SET "&msitablename&".Value='"&font&wert&"' WHERE "&msitablename&".Property='"&subtablename&"'"
									query = "UPDATE "&msitablename&" SET "&msitablename&".Message=? WHERE "&msitablename&".Error="&subtablename
								End if
								If tablename="Property" Then
									If LCase(subtablename)="language" Then wert=LCase(Wscript.Arguments(1))
									Record.StringData(1)=CStr(font&wert)
									'query = "UPDATE "&msitablename&" SET "&msitablename&".Value='"&font&wert&"' WHERE "&msitablename&".Property='"&subtablename&"'"
									query = "UPDATE "&msitablename&" SET "&msitablename&".Value=? WHERE "&msitablename&".Property='"&subtablename&"'"
								End If
								If tablename="UIText" Then										
									Record.StringData(1)=CStr(font&wert)
									query = "UPDATE `"&msitablename&"` SET `"&msitablename&"`.`Text`=? WHERE `"&msitablename&"`.`Key`='"&subtablename&"'"										
								end if
						    End If
					    end If
					    
						' MsgBox record.stringdata(1)&chr(13)&query
					    
				        Set view = database.OpenView(query) ':CheckError
				        'view.Execute : CheckError
				        View.Execute Record ': CheckError		
        		    end if 
				End If
  			Next
  			
  			feld=""
  			database.Commit:CheckError
		Next
		database.Commit:CheckError
		'deletefile folder&Wscript.Arguments(1)&".xml"
		'callprog("msiexec /I "&Wscript.Arguments(0))
	else
   		'msgbox "files do not exist"
	end if
else
	'msgbox "not enough arguments"
end If

Set view = nothing
Set database = nothing
Set installer =nothing

Sub CheckError
    If err.number<>0 Then
		err.clear
	End If
End Sub

sub deletefile(filename)
	on error resume next
	Set filesys = CreateObject("Scripting.FileSystemObject") 
	If filesys.FileExists(filename) Then
   		filesys.DeleteFile filename
	End If 
	set filesys=nothing
end sub


function folder
	set WshShell = WScript.CreateObject("WScript.Shell")
	folder=left(WScript.ScriptFullName,instrrev(WScript.ScriptFullName,"\"))	
	set WshShell=Nothing
end function

function checkfile(filename) 
	On error resume next
	Set filesys = CreateObject("Scripting.FileSystemObject") 
	If filesys.FileExists(filename) Then
   		checkfile=1
   	else 
   		checkfile=0
	End If 
	set filesys=nothing
end function

sub callprog(progname)
	Dim WshShell, msiinst
	Set WshShell = CreateObject("WScript.Shell")
	Set msiinst = WshShell.Exec(progname)
	WshShell.AppActivate msiinst.ProcessID
end sub

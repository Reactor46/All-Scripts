Dim strDomain, objFSO, objOutputFile
strDomain = wscript.arguments.item(0)
if strDomain = "" then
  	wscript.echo "Usage: cscript groupsmembers.vbs <OU Name>"
	wscript.echo "e.g. cscript ou=MyUsers,dc=wisesoft,dc=co,dc=uk"
  	wscript.quit
else
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objOutputFile = objFSO.CreateTextFile("groupsmembers.txt", 2, True)
	listOU strDomain

	objOutputFile.Close
	set objOutputFile = nothing
	set objFSO = nothing
end if 'usage

wscript.quit

function listOU(strOU)
	Dim objOU, objM
	set objOU=Getobject("LDAP://" & strOU)
	For each objM in objOU
  		Select case objM.class
    		Case "organizationalUnit"
      			listOU objM.distinguishedName
    		Case "group"
      			listgroup "", objM.distinguishedName
    		Case Else
      			'do nothing
  		End select
	Next
end function

function listgroup(strCrumb, strGroup)
	Dim objG, objM
	'wscript.echo "-- " & strGroup

	if strCrumb <> "" then strCrumb = strCrumb & "/"
	set objG = getObject("LDAP://" & strGroup)
	For each objM in objG.Members
  		Select case objM.class
    			Case "group"
      				if inStr(strCrumb, objM.cn) = 0 then 'prevent circular
        				listgroup strCrumb & objG.cn, objM.distinguishedName
      				end if
    			Case "user"
     				 wscript.echo strCrumb & objG.cn & "," & objM.cn & "," & objM.DisplayName
     				 objOutputFile.WriteLine strCrumb & objG.cn & "," & objM.cn & "," & objM.DisplayName
    			Case Else
      				wscript.echo "-------------" & objG.cn & " " & objM.cn & " " & objM.class
  		End select
	Next
end function
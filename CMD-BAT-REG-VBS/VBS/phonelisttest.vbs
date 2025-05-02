Option Explicit
on error resume next
Const ADS_SCOPE_SUBTREE = 2
Const ForAppending = 8
dim objconnection, objcommand, objrecordset, objfqdn, objname, nametruncate, objuser, objfso, strfileout, objfileout, objReportTextFile
dim wh, wsh, kal, ger, hen, ho, dwn, kta, bne, syd, dc, html1, html2, html3, ldapstring

Set objConnection = CreateObject("ADODB.Connection") 
Set objCommand =   CreateObject("ADODB.Command") 
objConnection.Provider = "ADsDSOObject" 
objConnection.Open "Active Directory Provider" 
Set objCommand.ActiveConnection = objConnection 
objCommand.Properties("Page Size") = 1000 
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCommand.Properties("Sort on") = "displayname"  'sort by display name
Set objFSO = CreateObject("Scripting.FileSystemObject")
ldapstring = "OU=users,DC=contoso,DC=network"

'set the output filename and open for editing
strFileOut = "path to file\phone-directory.htm"
Set objReportTextFile = objFSO.OpenTextFile(strFileOut, ForAppending, True)

'write the header information to the file
objReportTextFile.WriteLine("<html><head><title>Your company Phone Directory</title></head><body><br><a name=back> </a>")

'select names from ad
objCommand.CommandText = "SELECT Name FROM 'LDAP://" & ldapstring & "' "  

'set up variables for each site - this is so that you can store them on a per site/department basis. add/remove sites as you require
html1 = "<table frame=box><tr><td><a name="
html2 = "><h1>"
html3 = "</h1></a><a href=#back>Back to top</a> </td> <tr> <td> First Name </td><td> Surname </td><td> Phone </td><td> Title </td><td> mobile </td></tr>"
wsh = html1 & "wsh" & html2 & "Welshpool" & html3 
kal = html1 & "kal" & html2 & "Kalgoorlie" & html3
ger = html1 & "ger" & html2 & "Geraldton" & html3
hen = html1 & "hen" & html2 & "Henderson" & html3
ho = html1 & "ho" & html2 & "Headoffice" & html3
dwn = html1 & "dwn" & html2 & "Darwin" & html3
kta = html1 & "kta" & html2 & "Karratha" & html3
bne = html1 & "bne" & html2 & "Brisbane" & html3
syd = html1 & "syd" & html2 & "Sydney" & html3
dc = html1 & "dc" & html2 & "Distribution Centre" & html3

'table of contents
objfqdn = "<a href=#bne>Brisbane</a> <br> <a href=#dwn>Darwin</a> <br> <a href=#dc>Distribution Centre</a><br> <a href=#ger>Geraldton</a> <br> <a href=#hen>Henderson</a> <br> <a href=#ho>Headoffice</a> <br> <a href=#kal>Kalgoorlie</a> <br> <a href=#kta>Karratha</a> <br> <a href=#syd>Sydney</a> <br> <a href=#wsh>Welshpool</a> <br>"

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
objrecordset.movenext

Do Until objRecordSet.EOF
	objuser = ""
	objname = objRecordSet.Fields("Name").Value
	Set objUser = GetObject("LDAP://CN=" & objname & "," & ldapstring)
	nametruncate = left(objuser.givenName, 6)
	nametruncate = UCase(nametruncate) 'this is so we can do name tests later as we have generic accounts that should not be in the directory.
	if objUser.AccountExpirationDate = "1/01/1970" OR objUser.AccountExpirationDate = "1/01/1601 8:00:00 AM" Then 'look for expired accounts
		wh = 1
		generator("WELSHPOOL")
		wh = 666
		generator("CORPORATE OFFICE")
		wh = 6
		generator("HENDERSON")
		wh = 4
		generator("KALGOORLIE")
		wh = 3
		generator("GERALDTON")
		wh = 2
		generator("KARRATHA")
		wh = 7
		generator("DARWIN")
		wh = 5
		generator("BRISBANE")
		wh = 8
		generator("SYDNEY")
		wh = 9
		generator("DISTRIBUTION CENTRE")
		wh = 0
	else
		
	End If
	objRecordSet.MoveNext
Loop

'now that we've finished looping through all the accounts, close the tables and update the html file.
wsh = wsh & "</table><br>"
ho = ho & "</table><br>"
hen = hen & "</table><br>"
kal = kal & "</table><br>"
ger = ger & "</table><br>"
kta = kta & "</table><br>"
dwn = dwn & "</table><br>"
bne = bne & "</table><br>"
syd = syd & "</table><br>"
dc = dc & "</table><br>"
objreporttextfile.writeline(objfqdn & bne & dwn & dc & ger & hen & ho & kal & kta  & syd & wsh & "</body></html>" )


objreporttextfile.close

function generator(warehouse)

	if nametruncate = left(warehouse, 6) then
		
	else
		If left(nametruncate, 4) = "PORT" then 'this is a throwback to an older generic naming convention we used.
		
		Else
			If (objuser.telephoneNumber = "" And objuser.mobile = "") Or objuser.title = "" Then 'if they don't have a phone number, don't add them
			
			else 
				if Ucase(objuser.physicalDeliveryOfficeName) = warehouse then
					Select Case wh
						Case 1
							wsh = wsh & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 666
							ho = ho & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 6
							hen = hen & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 4
							kal = kal & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 3
							ger = ger & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 2
							kta = kta & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>"  & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 7
							dwn = dwn & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>"  & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 5
							bne = bne & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>"  & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 8
							syd = syd & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>" & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case 9
							dc = dc & "<tr> <td>" & objuser.givenName & "</td><td>" & objuser.sn & "</td><td>"  & objuser.telephoneNumber & "</td><td>" & objuser.title & "</td><td>" & objuser.mobile & "</td></tr>"
						Case else
					
					End Select		
				end if
			End If
		End If
	End If
End Function
set dlobj = createobject("CDO.Person")
dlobj.datasource.open "LDAP://DN",,3
Set iRecip = dlobj.GetInterface("IMailRecipient")
If iRecip.HideFromAddressBook = False then
	wscript.echo "Address List Currnetly Visable in GAL"
	Wscript.echo "Hiding Address List"
	iRecip.HideFromAddressBook = True
	dlobj.datasource.save
else
	wscript.echo "Address List Currnetly Hidden from GAL"
	Wscript.echo "Unhiding Address List"
	iRecip.HideFromAddressBook = False
	dlobj.datasource.save
end if
if iRecip.HideFromAddressBook = False then wscript.echo "Address List is Visable"
if iRecip.HideFromAddressBook = True then wscript.echo "Address List is Hidden"

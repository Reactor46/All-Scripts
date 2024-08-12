sUserADsPath = "LDAP://dn-of-addresslist"
Set objadlist = GetObject(sUserADsPath)
Set oSecurityDescriptor = objadlist.Get("ntSecurityDescriptor")
Set dacl = oSecurityDescriptor.DiscretionaryAcl
Set ace = CreateObject("AccessControlEntry")
wscript.echo "Here are the existing ACEs in the DACL:"
For Each ace In dacl
' Display all the properties of the ACEs using the IADsAccessControlEntry interface.
    wscript.echo ace.Trustee & ", " & ace.AccessMask & ", " & ace.AceType & ", " & ace.AceFlags & ", " & ace.Flags & ", " & ace.ObjectType & ", " & ace.InheritedObjectType
Next

wscript.echo "Done viewing descriptor"

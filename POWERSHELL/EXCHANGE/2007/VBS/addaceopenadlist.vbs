strDNofADdresslist = "LDAP://DN-of-Addresslist"
Set objgal = GetObject(strDNofADdresslist)
Set objSecurityDescriptor = objgal.Get("ntSecurityDescriptor")
iCurrentControl = objSecurityDescriptor.Control
objSecurityDescriptor.Control = iCurrentControl Or SE_DACL_PROTECTED
Set objParentSD = objgal.Get("ntSecurityDescriptor")
Set objParentDACL = objParentSD.DiscretionaryAcl
Set objCopyDACL = objParentDACL.CopyAccessList()

AddAce objCopyDACL, "domain\userID", 256, 5, 2, 1, "{A1990816-4298-11D1-ADE2-00C04FD8D5CD}", 0

Set objNewDACL = ReorderACL(objCopyDACL)
objSecurityDescriptor.DiscretionaryAcl = objNewDACL
objgal.Put "ntSecurityDescriptor", objSecurityDescriptor
objgal.SetInfo

Function AddAce( objDacl, _
                 szTrusteeName, _
                 gAccessMask, _
                 gAceType, _
                 gAceFlags, _
                 gFlags, _
                 gObjectType, _
                 gInheritedObjectType)
Set Ace1 = CreateObject("AccessControlEntry")
Ace1.AccessMask = gAccessMask
Ace1.AceType = gAceType
Ace1.AceFlags = gAceFlags
Ace1.Flags = gFlags
Ace1.Trustee = szTrusteeName
If CStr(gObjectType) <> "0" Then
	Ace1.ObjectType = gObjectType
End If
If CStr(gInheritedObjectType) <> "0" Then
	Ace1.InheritedObjectType = gInheritedObjectType
End If
objDacl.AddAce Ace1
Set Ace1 = Nothing
wscript.echo "Ace added"
End Function

Function ReorderACL(objDacl)
' Set Constants.
Const ADS_ACETYPE_ACCESS_DENIED = &H1
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const ADS_ACETYPE_ACCESS_ALLOWED = &H0
Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
Const ADS_ACEFLAG_INHERITED_ACE = &H10
Set objSD = CreateObject("SecurityDescriptor")
Set newDACL = CreateObject("AccessControlList")
Set ImpDenyDacl = CreateObject("AccessControlList")
Set ImpDenyObjectDacl = CreateObject("AccessControlList")
Set ImpAllowDacl = CreateObject("AccessControlList")
Set ImpAllowObjectDacl = CreateObject("AccessControlList")

For Each ace In objDacl
Select Case ace.AceType
      Case ADS_ACETYPE_ACCESS_DENIED
         ImpDenyDacl.AddAce ace
      Case ADS_ACETYPE_ACCESS_DENIED_OBJECT
         ImpDenyObjectDacl.AddAce ace
      Case ADS_ACETYPE_ACCESS_ALLOWED
         ImpAllowDacl.AddAce ace
      Case ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
         ImpAllowObjectDacl.AddAce ace
      Case Else
      	wscript.echo "bad ACE"
End Select
Next
    ' Combine the ACEs in the Proper Order
    ' Implicit Deny.
For Each ace In ImpDenyDacl
  newDACL.AddAce ace
Next
    ' Implicit Deny Object.
For Each ace In ImpDenyObjectDacl
  newDACL.AddAce ace
Next
   ' Implicit Allow.
For Each ace In ImpAllowDacl
  newDACL.AddAce ace
Next
' Implicit Allow Object.
For Each ace In ImpAllowObjectDacl
  newDACL.AddAce ace
Next
newDACL.AclRevision = objDacl.AclRevision
Set ReorderACL = newDACL
Set newDACL = Nothing
Set ImpAllowObjectDacl = Nothing
Set ImpAllowDacl = Nothing
Set ImpDenyObjectDacl = Nothing
Set ImpDenyDacl = Nothing
Set objSD = Nothing
wscript.echo "DACL Reordered"
End Function


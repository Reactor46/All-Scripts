dcDomainController = wscript.arguments(0)
set objIadsTools = CreateObject("IADsTools.DCFunctions")
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\AdrinfoLastupdated.csv",2,true) 
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Query = "<LDAP://" & strNameingContext & ">;(&(&(& (mailnickname=*)(mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) ))));distinguishedName,displayname;subtree"
Com.ActiveConnection = Conn
Com.CommandText = Query
Com.Properties("Page Size") = 1000
Set Rs = Com.Execute
wfile.writeline("Mailbox,TelephoneNumber,Mobile Number,Home Phone,Street Address,Office")
While Not Rs.EOF
	wfile.writeline(rs.fields("displayname").value & "," & getUserData(dcDomainController,rs.fields("distinguishedName")))
	rs.movenext
Wend
wfile.close
set wfile = nothing
set fso = Nothing



function getUserData(dcDomainController,dnUserDN)
dlDataline = ""
tnTelephoneNumber  = "Not Set"
mnMobileNumber  = "Not Set"
hnHomePhone = "Not Set"
saStreetAddress = "Not Set"
ofOffice = "Not Set"

intRes = objIadsTools.GetMetaData(Cstr(dcDomainController),Cstr(dnUserDN),0)

if intRes = -1 then
   Wscript.Echo objIadsTools.LastErrorText
   WScript.Quit
end if
wscript.echo "User" & dnUserDN
for count = 1 to intRes
   select case objIadsTools.MetaDataName(count) 
	case "telephoneNumber" wscript.echo "Telephone Number last write: " & objIadsTools.MetaDataLastWriteTime(count)
		      tnTelephoneNumber = objIadsTools.MetaDataLastWriteTime(count)
	case "mobile" wscript.echo "Mobile Number last write: " & objIadsTools.MetaDataLastWriteTime(count)
		      mnMobileNumber = objIadsTools.MetaDataLastWriteTime(count)	
	case "homePhone"  wscript.echo "Home Phone Number last write: " & objIadsTools.MetaDataLastWriteTime(count)
		      hnHomePhone = objIadsTools.MetaDataLastWriteTime(count)  
	case "streetAddress"  wscript.echo "Street Address last write: " & objIadsTools.MetaDataLastWriteTime(count)  
		      saStreetAddress = objIadsTools.MetaDataLastWriteTime(count) 
	case "physicalDeliveryOfficeName"  wscript.echo "Office last write: " & objIadsTools.MetaDataLastWriteTime(count)  
   		      ofOffice = objIadsTools.MetaDataLastWriteTime(count) 
   end select
next
wscript.echo 
dlDataline = tnTelephoneNumber & "," & mnMobileNumber & "," & hnHomePhone & "," & saStreetAddress  & "," & ofOffice
getUserData = dlDataline
end function


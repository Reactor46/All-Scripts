servername = wscript.arguments(0)
resultfile = "c:\oval\ovaltestresult.xml"
set objdom = CreateObject("MICROSOFT.XMLDOM")
Set objField = objDom.createElement("Exchange_Oval_Test_Results")
objDom.appendChild objField
Set objattID = objDom.createAttribute("Servername")
objattID.Text = servername
objField.setAttributeNode objattID
call RetrieveExchangeDefinition()
Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
objDom.insertBefore objPI, objDom.childNodes(0)
objdom.save(resultfile)
wscript.echo
wscript.echo "Scan Finished"

sub RetrieveExchangeDefinition()

set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
xdXmlDocument.async="false"
xdXmlDocument.load("c:\oval\windows.definitions.xml")
Set xnProductNodes = xdXmlDocument.selectNodes("//*[text() = 'Microsoft Exchange Server 2003']")
For Each xnProductNode In xnProductNodes 
  Set xnDefinitionNodes = xnProductNode.parentNode.parentNode.childNodes
  wscript.echo "Compliance Definition: " & xnProductNode.parentNode.parentNode.Attributes(0).Text
  Set objField1 = objDom.createElement(xnProductNode.parentNode.parentNode.Attributes(0).Text)
  objfield.appendChild objField1
  cpass = "Passed"
  For Each xnDefinitionNode In xnDefinitionNodes
  	select case xnDefinitionNode.NodeName
		case "description" wscript.echo "Description : " & xnDefinitionNode.text
				   Set objField2 = objDom.createElement("description")
				   objfield2.text = xnDefinitionNode.text & vbcrlf
				   objfield1.appendChild objField2
				   
		case "notes" wscript.echo "Notes : " & xnDefinitionNode.text
				   Set objField3 = objDom.createElement("Notes")
				   objfield3.text = xnDefinitionNode.text & vbcrlf
				   objfield1.appendChild objField3
		case "criteria" set xnCriteriaNodes = xnDefinitionNode.childNodes
		     		tpTesttoperform =  procCritera(xnCriteriaNodes)
				arTestarry = Split(tpTesttoperform, ";" )
				wscript.echo
				Wscript.echo "Number of Tests to be performed : " & ubound(arTestarry)
				Set objField4 = objDom.createElement("Tests")
				Set objattID = objDom.createAttribute("Number_Tests")
				objattID.Text = "" & ubound(arTestarry)
				objField4.setAttributeNode objattID
				objfield1.appendChild objField4
				wscript.echo
				for e = lbound(arTestarry) to ubound(arTestarry)-1
					Set xntestNodes = xdXmlDocument.selectNodes("//*[@id = '" & arTestarry(e) & "']")
					for each xnTestnode in xntestNodes
						Set objField5 = objDom.createElement(arTestarry(e))
						Set objattID = objDom.createAttribute("test_type")
						objattID.Text = xnTestnode.Nodename
						objField5.setAttributeNode objattID
						objfield4.appendChild objField5
						Wscript.echo "Test Going to be Proformed : " & xnTestnode.Nodename
						select case xnTestnode.Nodename
							case "registry_test"  testres =  registry_test(xnTestnode,objField5)
							case "activedirectory_test" testres = activedirectory_test(xnTestnode,objField5)
							case "file_test" testres = file_test(xnTestnode,objField5)
							case "wmi_test"  testres = wmi_test(xnTestnode,objField5)
							case "unknown_test" = unknown_test(xnTestnode,objField5)
						end select

					next
				next
	end select
  Next
  wscript.echo
Next
end sub

function procCritera(xnCriteriaNodes)
	tstarray = ""
	for each xnCriteriaTestnode in xnCriteriaNodes
		select case xnCriteriaTestnode.NodeName
			case "software" set xnSotfwareNodes = xnCriteriaTestnode.childNodes
					wscript.echo
					Wscript.echo "Software Dependenices and Negation"
					for each xnSoftwareNode in xnSotfwareNodes
					 	wscript.echo xnSoftwareNode.attributes.getNamedItem("comment").nodeValue _
							& " " & xnSoftwareNode.attributes.getNamedItem("negate").nodeValue
						tstarray = tstarray & xnSoftwareNode.attributes.getNamedItem("test_ref").nodeValue & ";"
					next 
			case "configuration" set xnconfigNodes = xnCriteriaTestnode.childNodes
					Wscript.echo
					Wscript.echo "Software Configuration and Negation"
					for each xnconfigNode in xnconfigNodes
					 	wscript.echo xnconfigNode.attributes.getNamedItem("comment").nodeValue & " " _
							 & xnconfigNode.attributes.getNamedItem("negate").nodeValue
						tstarray = tstarray & xnconfigNode.attributes.getNamedItem("test_ref").nodeValue & ";"
					next 
		end select
	next
	procCritera = tstarray
end function

function registry_test(Testdef,objField5)
wscript.echo
wscript.echo testdef.attributes.getNamedItem("comment").nodeValue
wscript.echo
Set objattID = objDom.createAttribute("comment")
objattID.Text = testdef.attributes.getNamedItem("comment").nodeValue
objField5.setAttributeNode objattID
set ndregnodes = testdef.childNodes
for each ndregnode in ndregnodes
	set ndrerkeynodes = ndregnode.childNodes
	for each ndrerkeynode in ndrerkeynodes
		Set objField6 = objDom.createElement(ndrerkeynode.nodename)
		objfield6.text = ndrerkeynode.text & vbcrlf
		objfield5.appendChild objField6
		wscript.echo ndrerkeynode.nodename & " : " & ndrerkeynode.text
		select case ndrerkeynode.nodename 
			case "hive" hiQueryhive = ndrerkeynode.text
			case "key" kyQueryky = ndrerkeynode.text
			case "name" nmQueryname = ndrerkeynode.text
			case "value" vlQueryvalue = ndrerkeynode.text
		end select
	next
next
Wscript.echo "Performing Test"
regval	= ""
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
& servername & "\root\default:StdRegProv")
objReg.EnumValues &H80000002, kyQueryky, arValues, arvaluetypes
if isnull(arValues) then
 	Wscript.echo 
	wscript.echo "Key Not Found"
	registry_test = "Fail"
	wscript.echo
	set objField7 = objDom.createElement("Result")
	objfield7.text = "Key Not Found" & vbcrlf
	objfield5.appendChild objField7
else
	Wscript.echo 
	wscript.echo "Key Found"
	if nmQueryname = ".*" then
		wscript.echo "Test Passed"
		set objField7 = objDom.createElement("Result")
		objfield7.text = "Key Found" & vbcrlf
		objfield5.appendChild objField7
	else
		for vl = lbound(arValues) to ubound(arValues)
			if arValues(vl) = nmQueryname then 
				wscript.echo arValues(vl) & " " & arvaluetypes(vl)
				select case arvaluetypes(vl)
					case 1 objReg.GetStringValue &H80000002, kyQueryky, nmQueryname, regval
					case 2 objReg.GetExpandedStringValue &H80000002, kyQueryky, nmQueryname, regval
					case 3 objReg.GetBinaryvalue &H80000002, kyQueryky, nmQueryname, regval
					case 4 objReg.GetDwordvalue &H80000002, kyQueryky, nmQueryname, regval
					case 7 objReg.Getmultistringvalue &H80000002, kyQueryky, nmQueryname, regval
				end select
				set objField7 = objDom.createElement("Result")
				objfield7.text = regval
				objfield5.appendChild objField7
			end if	
		next
	end if
	wscript.echo
end if
end function

function file_test(Testdef,objField5)
fversion = ""
wscript.echo
wscript.echo testdef.attributes.getNamedItem("comment").nodeValue
wscript.echo
Set objattID = objDom.createAttribute("comment")
objattID.Text = testdef.attributes.getNamedItem("comment").nodeValue
objField5.setAttributeNode objattID
set ndfilenodes = testdef.childNodes
for each ndfilenode in ndfilenodes
	set ndrerkeynodes = ndfilenode.childNodes
	for each ndrerkeynode in ndrerkeynodes
		set ndtypekeynodes = ndrerkeynode.childNodes
		for each ndtypekeynode in ndtypekeynodes
			if ndtypekeynode.nodename = "component" then 
				if ndtypekeynode.attributes.getNamedItem("type").nodeValue = "registry_value" then regpathfile = ndtypekeynode.text 
				if ndtypekeynode.attributes.getNamedItem("type").nodeValue = "literal" then literalfile = ndtypekeynode.text 
			else  
				if ndtypekeynode.parentNode.nodename = "version" then
					if fversion = "" then
						fversion = ndtypekeynode.text 
					else
						fversion = fversion & "." & ndtypekeynode.text 
					end if
				end if
			end if
		next
	next
next
Wscript.echo "Performing Test"
wscript.echo
rkTestRegkey = ""
arRegkeyarry = Split(regpathfile, "\" )
for rk = lbound(arRegkeyarry)+1 to ubound(arRegkeyarry)-1
	 if rkTestRegkey = "" then
	 	rkTestRegkey = arRegkeyarry(rk)
	 else
		rkTestRegkey = rkTestRegkey & "\" & arRegkeyarry(rk)
	 end if 
next
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
& servername & "\root\default:StdRegProv")
objReg.GetStringValue &H80000002, rkTestRegkey, arRegkeyarry(ubound(arRegkeyarry)), exfileregval

wscript.echo rkTestRegkey
wscript.echo arRegkeyarry(ubound(arRegkeyarry))
wscript.echo exfileregval
wscript.echo literalfile
wscript.echo fversion
wscript.echo

Set objField6 = objDom.createElement("FileName")
objfield6.text = exfileregval & literalfile  & vbcrlf
objfield5.appendChild objField6
Set objField7 = objDom.createElement("Lowest_Version")
objfield7.text = fversion  & vbcrlf
objfield5.appendChild objField7

Set efile = GetObject("winMgmts:!\\" & servername & "\root\cimv2:CIM_DataFile.Name='" & exfileregval & literalfile & "'")
wscript.echo "Actual File Version: " & efile.Version
Set objField8 = objDom.createElement("Result")
objfield8.text = efile.Version  & vbcrlf
objfield5.appendChild objField8
if efile.Version > fversion then
	wscript.echo "Is greator then required"
	tlTestlog = tlTestlog & "Is greator then required" & vbcrlf & vbcrlf
else
	wscript.echo "Is Less then required"
	tlTestlog = tlTestlog & "Is Less then required" & vbcrlf & vbcrlf
end if
end function

function activedirectory_test(Testdef,objField5)

wscript.echo
wscript.echo testdef.attributes.getNamedItem("comment").nodeValue
Set objattID = objDom.createAttribute("comment")
objattID.Text = testdef.attributes.getNamedItem("comment").nodeValue
objField5.setAttributeNode objattID
wscript.echo
set ndactdirnodes = testdef.childNodes
for each ndactdirnode in ndactdirnodes
	set ndctdirkeynodes = ndactdirnode.childNodes
		for each ndactdirkeynode in ndctdirkeynodes
			select case ndactdirkeynode.parentnode.nodename 
				case "object" 	select case ndactdirkeynode.nodename
							case "naming_context" ncNamingcontext = ndactdirkeynode.text
							case "relative_dn" dnRelativeDN = ndactdirkeynode.text
							case "attribute" atAdattribute = ndactdirkeynode.text
						end select					
				case "data"     if ndactdirkeynode.nodename = "value" then
					       		advalop = ndactdirkeynode.attributes.getNamedItem("operator").nodeValue  
					    		advalue = ndactdirkeynode.text
						end if
			end select
		next

next
Wscript.echo ncNamingcontext
wscript.echo dnRelativeDN
wscript.echo atAdattribute
wscript.echo advalop
wscript.echo advalue
wscript.echo

Rem Oval Def Fixups
if mid(dnRelativeDN,len(dnRelativeDN),1) = "$" then
	dnRelativeDN = mid(dnRelativeDN,1,len(dnRelativeDN)-1)
end if
dnRelativeDN = replace(dnRelativeDN,"^Exchange,","^CN=Exchange,")
dnRelativeDN = replace(dnRelativeDN,"^Public,","^CN=Public,")
dnRelativeDN = replace(dnRelativeDN,"[^,]+","[^,]*")
dnRelativeDN = replace(dnRelativeDN,"attribute>","")
wscript.echo dnRelativeDN
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
qserver = 0
if instr(dnRelativeDN,",CN=Servers,") then
	qserver = 1
	svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName;subtree"
	Com.ActiveConnection = Conn
	Com.CommandText = svcQuery
	Set snrs = Com.Execute
	wscript.echo snrs.fields("distinguishedName")
end if
select case ncNamingcontext
	case "configuration" strNameingContext = iAdRootDSE.Get("configurationNamingContext")	
	case "default"  strNameingContext = iAdRootDSE.Get("defaultNamingContext")	
end select
atrQuery = "<LDAP://" & strNameingContext & ">;(&(objectclass=*)(" & atAdattribute & "=*));cn,name,distinguishedName;subtree"
wscript.echo atrQuery 
Com.ActiveConnection = Conn
Com.CommandText = atrQuery
Set Rs = Com.Execute
Set oRE = New RegExp
oRE.IgnoreCase = True
oRE.Pattern = dnRelativeDN
wscript.echo "ADSI queried Values"
wscript.echo
dnDNstotest = ""
While Not Rs.EOF
	wscript.echo rs.fields("distinguishedName").value
	bMatch = oRE.Test(rs.fields("distinguishedName").value)
	Wscript.echo "RegEX test result : " & bMatch
	if bMatch = true then
		if qserver = 1 then
			if instr(rs.fields("distinguishedName").value,snrs.fields("distinguishedName")) then
				dnDNstotest = dnDNstotest & rs.fields("distinguishedName").value & ";"
				wscript.echo "Found Server Match"
			end if
		else 
			dnDNstotest = dnDNstotest & rs.fields("distinguishedName").value  & ";"	
		end if
	end if
	rs.movenext
wend
if dnDNstotest = "" then
	set objField8 = objDom.createElement("Attribute")
	objfield8.text = atAdattribute & vbcrlf
	objfield5.appendChild objField8
	set objField9 = objDom.createElement("result")
	objfield9.text = "Property not set in Active Directory" & vbcrlf
	objfield5.appendChild objField9
else
	Wscript.echo
	Wscript.echo "DN's to be tested"
	wscript.echo
	arrydnstotest = Split(dnDNstotest, ";" )
	for ads = lbound(arrydnstotest) to ubound(arrydnstotest)-1
		set objField7 = objDom.createElement("AD_Distinguished_Name")
		objfield7.text = arrydnstotest(ads) & vbcrlf
		objfield5.appendChild objField7
		set objField8 = objDom.createElement("Attribute")
		objfield8.text = atAdattribute & vbcrlf
		objfield5.appendChild objField8
		set objField8b = objDom.createElement("Operator")
		objfield8b.text = advalop & vbcrlf
		objfield5.appendChild objField8b
		set objField9 = objDom.createElement("Value")
		objfield9.text = advalue & vbcrlf
		objfield5.appendChild objField9
		dntoget = ""
		dntoget = "LDAP://" & arrydnstotest(ads)
		wscript.echo dntoget
		set aoADobject = getobject(dntoget)
		' ---------------------------------------------------------------
		' Type checking From the book "Active Directory Cookbook" by Robbie Allen
		' http://www.rallenhome.com/books/adcookbook/src/10.08-view_attribute.vbs.txt
		' ---------------------------------------------------------------

		set dicADsType = CreateObject("Scripting.Dictionary")
		dicADsType.Add 0, "INVALID"
		dicADsType.Add 1, "DN_STRING"
		dicADsType.Add 2, "CASE_EXACT_STRING"
		dicADsType.Add 3, "CASE_IGNORE_STRING"
		dicADsType.Add 4, "PRINTABLE_STRING"
		dicADsType.Add 5, "NUMERIC_STRING"
		dicADsType.Add 6, "BOOLEAN"
		dicADsType.Add 7, "INTEGER"
		dicADsType.Add 8, "OCTET_STRING"
		dicADsType.Add 9, "UTC_TIME"
		dicADsType.Add 10, "LARGE_INTEGER"
		dicADsType.Add 11, "PROV_SPECIFIC"
		dicADsType.Add 12, "OBJECT_CLASS"
		dicADsType.Add 13, "CASEIGNORE_LIST"
		dicADsType.Add 14, "OCTET_LIST"
		dicADsType.Add 15, "PATH"
		dicADsType.Add 16, "POSTALADDRESS"
		dicADsType.Add 17, "TIMESTAMP"
		dicADsType.Add 18, "BACKLINK"
		dicADsType.Add 19, "TYPEDNAME"
		dicADsType.Add 20, "HOLD"
		dicADsType.Add 21, "NETADDRESS"
		dicADsType.Add 22, "REPLICAPOINTER"
		dicADsType.Add 23, "FAXNUMBER"
		dicADsType.Add 24, "EMAIL"
		dicADsType.Add 25, "NT_SECURITY_DESCRIPTOR"
		dicADsType.Add 26, "UNKNOWN"
		dicADsType.Add 27, "DN_WITH_BINARY"
		dicADsType.Add 28, "DN_WITH_STRING"
		set  objObject = getobject(dntoget)
		objObject.GetInfo
   		set objPropEntry = objObject.Item(atAdattribute)
		flagfail = ""
		flagpass = ""
   		for Each objPropValue In objPropEntry.Values  
      			adqvalue = ""  
		        if (dicADsType(objPropValue.ADsType) = "DN_STRING") then
            			adqvalue = objPropValue.DNString
		        elseIf (dicADsType(objPropValue.ADsType) = "CASE_EXACT_STRING") then
            			adqvalue = objPropValue.CaseExactString
         		elseIf (dicADsType(objPropValue.ADsType) = "CASE_IGNORE_STRING") then
            			adqvalue = objPropValue.CaseIgnoreString
		        elseIf (dicADsType(objPropValue.ADsType) = "PRINTABLE_STRING") then
            			adqvalue = objPropValue.PrintableString
		        elseIf (dicADsType(objPropValue.ADsType) = "NUMERIC_STRING") then
            			adqvalue = objPropValue.NumericString
		        elseIf (dicADsType(objPropValue.ADsType) = "BOOLEAN") then
            			adqvalue = CStr(objPropValue.Boolean)
		        elseIf (dicADsType(objPropValue.ADsType) = "INTEGER") then
            			adqvalue = objPropValue.Integer
		        elseIf (dicADsType(objPropValue.ADsType) = "LARGE_INTEGER") then
            		set objLargeInt = objPropValue.LargeInteger
            			adqvalue = objLargeInt.HighPart * 2^32 + objLargeInt.LowPart
				wscript.echo "Large Int"
		        elseIf (dicADsType(objPropValue.ADsType) = "UTC_TIME") then
		          	adqvalue = objPropValue.UTCTime
		        else
            			adqvalue =  dicADsType.Item(objPropEntry.ADsType) 
		        end if
			WScript.Echo objPropEntry.Name & " : " & adqvalue
			set objField10 = objDom.createElement("Result")
			objField10.text = adqvalue & vbcrlf
			objfield5.appendChild objField10
			if advalop = "bitwise and" then
				if adqvalue < 2147483647 and adqvalue > -2147483647 and advalue < 2147483647 and advalue > -2147483647  then
					if adqvalue and advalue then
						bitwise = "Bit Set"
					else 
						bitwise = "Bit Not Set"
					end if
				else
					bitwise = "Value to Large to test with VBS"
			        end if
				set objField11 = objDom.createElement("BitWise_Test")
				objField11.text = bitwise & vbcrlf
				objfield5.appendChild objField11
			end if
		next
	Next
end if
end function

function wmi_test(Testdef,objField5)
on error resume next
wscript.echo
wscript.echo testdef.attributes.getNamedItem("comment").nodeValue
Set objattID = objDom.createAttribute("comment")
objattID.Text = testdef.attributes.getNamedItem("comment").nodeValue
objField5.setAttributeNode objattID
wscript.echo

set ndwmidirnodes = testdef.childNodes
for each ndwmidirnode in ndwmidirnodes
	set ndctdirkeynodes = ndwmidirnode.childNodes
		for each ndwmidirkeynode in ndctdirkeynodes
			select case ndwmidirkeynode.parentnode.nodename 
				case "object" 	select case ndwmidirkeynode.nodename
							case "namespace" ncwmiNamespace = ndwmidirkeynode.text
									ncwmiop =   ndwmidirkeynode.attributes.getNamedItem("operator").nodeValue
							case "wql" wmiwql = ndwmidirkeynode.text
								 	wqlwmiop =  ndwmidirkeynode.attributes.getNamedItem("operator").nodeValue
						end select					
			end select
		next
next
set objField6 = objDom.createElement("WMINamespace")
objfield6.text = ncwmiNamespace & vbcrlf
objfield5.appendChild objField6
set objField7 = objDom.createElement("WQL")
objfield7.text = wmiwql & vbcrlf
objfield5.appendChild objField7
wscript.echo ncwmiNamespace
wscript.echo ncwmiop
wscript.echo wmiwql
wscript.echo wqlwmiop
wscript.echo
wscript.echo "Performing WMI Query"
wscript.echo 
wscript.echo "WMI Results"
wscript.echo
if instr(wmiwql,"select ") then wmiwql = replace(wmiwql,"select ","select name,") 
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername & "/" & ncwmiNamespace
Set objWMI =  GetObject(strWinMgmts)
Set listWMInodes = objWMI.ExecQuery(wmiwql)
For each WMInode in listWMInodes
	if err.number <> 0 then
		set objField8 = objDom.createElement("Result")
		objfield8.text = "Nothing Found Test only supported on Windows 2003" & vbcrlf
		objfield5.appendChild objField8
	else
		wscript.echo WMInode.name
		set objField8 = objDom.createElement("Result")
		objfield8.text = WMInode.name & vbcrlf
		objfield5.appendChild objField8
	end if
next
wscirpt.echo
wmi_test = "Unknown"
on error goto 0
end function

function unknown_test(Testdef,objField5)

wscript.echo
wscript.echo testdef.attributes.getNamedItem("comment").nodeValue
Set objattID = objDom.createAttribute("comment")
objattID.Text = testdef.attributes.getNamedItem("comment").nodeValue
objField5.setAttributeNode objattID
wscript.echo

end function
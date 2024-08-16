<%
Set oForm = Server.CreateObject("WSS.Form")
dim sysinfo
dim oUser
Set sysinfo = CreateObject("ADSystemInfo")
set Dirobj = GetObject("LDAP:")
Set oUser = Dirobj.OpenDSObject("LDAP://DCservername/"& sysinfo.UserName, "domain\user", "password", 0)
If Request.ServerVariables("REQUEST_METHOD") = "GET"  Then
	oForm.fields("FName").value = oUser.GivenName
	oForm.fields("LName").value = oUser.sn
	oForm.fields("Bphone").value = oUser.TelephoneNumber
	oForm.fields("Mphone").value = oUser.mobile
	oForm.fields("Hphone").value = oUser.homephone
	oForm.fields.update
 	oForm.Render
else
   If Request.ServerVariables("REQUEST_METHOD") = "POST"  Then
    	if oForm.Fields("FName").Value = "" then
	    oForm.Elements("FName").ErrorString = "<span style=""COLOR:red"">First Name Required<span>"        
            oForm.Render
	elseif oForm.Fields("LName").Value = "" then
	    oForm.Elements("LName").ErrorString = "<span style=""COLOR:red"">Last Name Required<span>"         
            oForm.Render
        else
		oUser.GivenName =  oForm.Fields("FName").Value
		oUser.sn = oForm.Fields("LName").Value
		oUser.TelephoneNumber = oForm.Fields("BPhone").Value
		if oForm.Fields("BPhone").Value = "" then
			oUser.putex 1,"TelephoneNumber", vbNull
		else
			oUser.TelephoneNumber = oForm.Fields("BPhone").Value
		end if 	 
		if oForm.Fields("MPhone").Value = "" then
			oUser.putex 1,"mobile", vbNull
		else
			oUser.mobile = oForm.Fields("MPhone").Value
		end if 	  
		if oForm.Fields("HPhone").Value = "" then
			oUser.putex 1,"homephone", vbNull
		else
			oUser.homephone = oForm.Fields("HPhone").Value
		end if 	 
		oUser.setinfo
		oForm.Elements("Result").ErrorString = "<span style=""COLOR:red"">Updated<span>"         
            	oForm.Render
	 end if  
    else 
    end if   
end if 
%>

<HTML>
<HEAD>
<BASE TARGET="_top">

</HEAD>

<BODY><!The Data URL macro is expanded at runtime by
the renderer so that the Submit
button will post back to the item itself.>
<FORM action="" id=FORM1 method=post name="FORM1"
target="_self">

<H1> Phone Number Details 
<INPUT class="field" name="result"
style="HEIGHT: 25px; WIDTH: 0px"></H1>
<br><br>

<b>First Name:</b> <INPUT class="field" name="FName"
style="HEIGHT: 25px; WIDTH: 200px"> <br><br>
<b>Last Name:</b> <INPUT class="field" name="LName"
style="HEIGHT: 25px; WIDTH: 200px"> <br><br>
<b>Business PhoneNumber:</b> <INPUT class="field"
name="Bphone" style="HEIGHT: 23px; WIDTH: 200px">
<br><br>
<b>Mobile PhoneNumber:</b><INPUT class="field"
name="Mphone" style="HEIGHT: 23px; WIDTH: 200px">
<br><br>
<b>Home Phone:</b> <INPUT class="field" name="Hphone"
style="HEIGHT: 25px; WIDTH: 200px"> <br>
<br><br>
&nbsp;<INPUT id=submit1 name=submit1 type=submit
value=Submit>

</FORM>

</BODY>
</HTML>

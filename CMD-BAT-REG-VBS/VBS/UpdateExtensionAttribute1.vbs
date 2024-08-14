'Use this script to update extensionAttribute1 with the user profile picture URLs.

‘Create the Variables that will be needed in the script
Dim s
Dim PictureURL
Const ADS_PROPERTY_CLEAR = 1
Dim PhotoLibrary = “http://my.contoso.com/employeephotos/”
SET objParent = GETOBJECT(LDAP://OU=Users,DC=Contoso,DC=Com)
objparent.FILTER = ARRAY("user")

 
With CreateObject("MSXML2.XMLHTTP")
                 
‘Loop through each user in the specified OU
FOR EACH objUser in objParent
 
'Notify Administrators Which Object Is Being Evaluated
 Wscript.Echo "Evaluating " & objUser.Get("CN")
 
'Set the URL to the user photo location
     PictureURL =  PhotoLibrary & objUser.GET("SAMAccountName") & ".jpg"
 
'Retrieve the header for the document
     .open "HEAD", PictureURL, false
     .send
     s = .status
     
    ‘If we retrieve a HTTP status of 200 (OK), the picture for the user exists.  Update the Extension Attribute
    if(s = 200) then 
        objUser.put "extensionAttribute1”, PictureURL
         objuser.Setinfo    
        Wscript.Echo("Photo for user " &Objuser.GET("CN") & “ exists.  ExtensionAttribute1 has been set to “ &PictureURL)
     elseif(s = 404) then 
        Wscript.Echo("Photo For User " &Objuser.GET("CN") & " does not exist.  ExtensionAttribute1 has been cleared.")
     objUser.putex ADS_PROPERTY_CLEAR, "extensionAttribute1", 0
         objuser.Setinfo 
    else 
       Wscript.Echo(s)  
    end if
 NEXT
 End With

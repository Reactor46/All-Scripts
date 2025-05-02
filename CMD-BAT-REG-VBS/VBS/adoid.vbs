OWAURL = "http://server/public/folder"
set conn = createobject("ADODB.Connection")
Conn.Provider = "ExOLEDB.DataSource"
conn.open OWAURL
set rec = createobject("ADODB.Record")
rec.open OWAURL,conn,3
wscript.echo Octenttohex(rec.fields("http://schemas.microsoft.com/mapi/proptag/x0FFF0102")) 



Function Octenttohex(OctenArry)  
  ReDim aOut(UBound(OctenArry)) 
  For i = 1 to UBound(OctenArry) + 1 
    if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
    	aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
    else
	aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
    end if
  Next 
  Octenttohex = join(aOUt,"")
End Function 

Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 

NombredelaLista="Poblacion"
URLColeccion="http://.../_vti_bin/"

'Crea la lista TemplateID=100 Lista Generica

Peticion="<?xml version='1.0' encoding='utf-8'?>" + _
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>" + _
         "<soap:Body>" + _
		 "<AddList xmlns='http://schemas.microsoft.com/sharepoint/soap/'>" + _
         "<listName>"+NombredelaLista+"</listName>" + _
         "<description>Cifras de población resultantes de la Revisión del Padrón municipal a 1 de enero de 2011. Datos de www.ine.es</description>" + _
         "<templateID>100</templateID>" + _
         "</AddList>" + _
         "</soap:Body>" + _
		 "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/AddList"
URLServicio = URLColeccion+"lists.asmx"

EjecutaPeticion

'Modifica la Lista, renombrando el campo Title y agregando los campos nuevos

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<UpdateList xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
         "<listName>"+NombredelaLista+"</listName>"+_
         "<newFields>"+_
		 "<Fields>"+_
	     "<Method ID='1'>"+_
		 "<Field Type='Text' DisplayName='CPRO' Description='Código de Provincia' Required='TRUE'/>"+_
	     "</Method>"+_
	     "<Method ID='2'>"+_
		 "<Field Type='Text' DisplayName='Provincia' Description='Provincia' Required='TRUE'/>"+_
	     "</Method>"+_
	     "<Method ID='3'>"+_
		 "<Field Type='Text' DisplayName='CMUN' Description='Código de Municipio' Required='TRUE'/>"+_
	     "</Method>"+_
	     "<Method ID='4'>"+_
		 "<Field Type='Text' DisplayName='Municipio' Description='Municipio' Required='TRUE'/>"+_
	     "</Method>"+_
	     "<Method ID='5'>"+_
    	 "<Field Type='Number' DisplayName='Varones' Description='Cantidad de Varones'/>"+_
	     "</Method>"+_
	     "<Method ID='6'>"+_
		 "<Field Type='Number' DisplayName='Mujeres' Description='Cantidad de Mujeres'/>"+_
	     "</Method>"+_
		 "<Method ID='7'>"+_
         "<Field Type='Calculated' DisplayName='Total' ResultType='Number' Description='Total Población'>"+_
         "<Formula>=Varones+Mujeres</Formula>"+_
         "</Field>"+_
         "</Method>"+_
		 "</Fields>"+_
     	 "</newFields>"+_
		 "<updateFields>"+_
		 "<Fields>"+_
		 "<Method ID='8'>"+_
		 "<Field Type='Text' Name='Title' DisplayName='Comunidad' Description='Comunidad Autónoma' Required='TRUE'/>"+_
		 "</Method>"+_
		 "</Fields>"+_
		 "</updateFields>"+_
         "</UpdateList>"+_
         "</soap:Body>"+_
         "</soap:Envelope>"

AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/UpdateList"

EjecutaPeticion

'Lee el archivo DatosPoblacion.txt separado por VTAB en lotes y los insertamos en la lista creada
Lote=100
Contador=0
Batch="<Batch OnError='Continue'>"

AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"

Set Fichero = ObjetoArchivo.OpenTextFile("DatosPoblacion.txt", 1) 
Do While Fichero.AtEndOfStream <> True
    Batch=Batch+ Divide(Fichero.ReadLine)
	Contador=Contador+1
	If (Contador Mod Lote)=0 Then
		Batch=Batch+"</Batch>"
		InsertaBatch
		Batch="<Batch OnError='Continue'>"
	End If
Loop
Fichero.Close
'Inserta los registros restantes
Batch=Batch+"</Batch>"
InsertaBatch

'Modifica la vista Allitems creada por defecto para que muestre las columnas agregadas

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<UpdateView xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
         "<listName>"+NombredelaLista+"</listName>"+_
         "<viewFields><ViewFields>"+_
         "<FieldRef Name='Title'/><FieldRef Name='CPRO'/>"+_
         "<FieldRef Name='Provincia'/><FieldRef Name='CMUN'/>"+_
         "<FieldRef Name='Municipio'/><FieldRef Name='Varones'/>"+_
         "<FieldRef Name='Mujeres'/><FieldRef Name='Total'/>"+_
         "</ViewFields></viewFields>"+_
         "</UpdateView>"+_
         "</soap:Body>"+_
         "</soap:Envelope>"


AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/UpdateView"
URLServicio = URLColeccion+"Views.asmx"

EjecutaPeticion

'Crea una vista MasPoblados con los municipios mayores a 100.000 habitantes

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<AddView xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
         "<listName>"+NombredelaLista+"</listName>"+_
         "<viewName>MasPoblados</viewName>"+_
         "<viewFields><ViewFields>"+_
         "<FieldRef Name='Title'/><FieldRef Name='Provincia'/>"+_
         "<FieldRef Name='Municipio'/><FieldRef Name='Varones'/>"+_
         "<FieldRef Name='Mujeres'/><FieldRef Name='Total'/>"+_
         "</ViewFields></viewFields>"+_
         "<query><Query>"+_
         "<Where><Gt><FieldRef Name='Total' /><Value Type='Number'>100000</Value></Gt></Where><OrderBy><FieldRef Name='Total' Ascending='FALSE' /></OrderBy>"+_
         "</Query></query>"+_
         "</AddView>"+_
         "</soap:Body>"+_
         "</soap:Envelope>"

AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/AddView"

EjecutaPeticion

'Crea una vista Agrupados agrupando por Comunidad y Provincia para obtener la Suma de los Varones y Mujeres

'Crea la vista con los campos

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<AddView xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
         "<listName>"+NombredelaLista+"</listName>"+_
         "<viewName>Agrupados</viewName>"+_
         "<viewFields><ViewFields>"+_
         "<FieldRef Name='Title'/><FieldRef Name='Provincia'/>"+_
         "<FieldRef Name='Municipio'/><FieldRef Name='Varones'/>"+_
         "<FieldRef Name='Mujeres'/><FieldRef Name='Total'/>"+_
         "</ViewFields></viewFields>"+_
		 "<query></query>"+_
         "</AddView>"+_
         "</soap:Body>"+_
         "</soap:Envelope>"

EjecutaPeticion

'Recupera el ID de la Vista Creada
ObjetoDOM.loadXML(ObjetoHttp.responseText)
ID = ObjetoDOM.selectSingleNode("//View").getAttribute("Name")

'Actualiza la vista agrupando y sumando. Será la vista por defecto
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<UpdateView xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
         "<listName>"+NombredelaLista+"</listName>"+_
         "<viewName>"+ID+"</viewName>"+_
		 "<viewProperties><View DefaultView='TRUE' ></View></viewProperties>"+_
         "<query><Query><GroupBy Collapse='TRUE'><FieldRef Name='Title'/><FieldRef Name='Provincia'/></GroupBy></Query></query>"+_
         "<aggregations><Aggregations Value='On'>"+_
         "<FieldRef Name='Municipio' Type='COUNT' /><FieldRef Name='Varones' Type='SUM' /><FieldRef Name='Mujeres' Type='SUM' />"+_
         "</Aggregations></aggregations>"+_
         "</UpdateView>"+_
         "</soap:Body>"+_
         "</soap:Envelope>"

AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/UpdateView"

EjecutaPeticion


WScript.Echo "Fin del Proceso"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Function Divide(Texto)
Vector = Split(Texto, vbTab)
Divide = "<Method ID='"&Contador&"' Cmd='New'>"
Divide = Divide +"<Field Name='Title'>"+Vector(0)+"</Field>"
Divide = Divide +"<Field Name='CPRO'>"+Vector(1)+"</Field>"
Divide = Divide +"<Field Name='Provincia'>"+Vector(2)+"</Field>"
Divide = Divide +"<Field Name='CMUN'>"+Vector(3)+"</Field>"
Divide = Divide +"<Field Name='Municipio'>"+Vector(4)+"</Field>"
Divide = Divide +"<Field Name='Varones'>"+Vector(5)+"</Field>"
Divide = Divide +"<Field Name='Mujeres'>"+Vector(6)+"</Field>"
Divide = Divide +"</Method>"
End Function

Private Sub InsertaBatch
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
	  	 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
		 "<soap:Body>"+_
		 "<UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
		 "<listName>"+NombredelaLista+"</listName>"+_
		 "<updates>"+Batch+"</updates>"+_
		 "</UpdateListItems>"+_
		 "</soap:Body>"+_
		 "</soap:Envelope>"
		 
EjecutaPeticion

End Sub


  
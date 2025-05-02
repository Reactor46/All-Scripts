
Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 

'Petición de los grupos de la coleccion de sitios
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/directory/GetGroupCollectionFromSite"
URLServicio = "http://.../_vti_bin/UserGroup.asmx"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<GetGroupCollectionFromSite xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/' />"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

'Lectura del nombre de cada uno de los grupos
ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Grupos = ObjetoDOM.getElementsByTagName("Group") 

Set Fichero = ObjetoArchivo.CreateTextFile("SalidaGruposyUsuarios.txt", True) 
Fichero.WriteLine ("Nombre del Grupo"+VBTab+"Nombre del Usuario"+VBTab+"Login"+VBTab+"Correo"+VBTab+"Es Grupo")

For i=0 To Grupos.length-1
	NombreGrupo = Grupos.item(i).getAttribute("Name")
	Fichero.WriteLine (NombreGrupo)
	'Petición de los integrantes de cada grupo
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/directory/GetUserCollectionFromGroup"
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
	"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
	"<soap:Body>"+_
	"<GetUserCollectionFromGroup xmlns='http://schemas.microsoft.com/sharepoint/soap/directory/'>"+_
	"<groupName>"+NombreGrupo+"</groupName>"+_
	"</GetUserCollectionFromGroup>"+_
	"</soap:Body>"+_
	"</soap:Envelope>"
	EjecutaPeticion
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set Usuarios = ObjetoDOM.getElementsByTagName("User")
	'Lectura de las propiedades de cada usuario
	For j=0 To Usuarios.length-1
		Nombre = Usuarios.item(j).getAttribute("Name")
		Login = Usuarios.item(j).getAttribute("LoginName")
		Correo = Usuarios.item(j).getAttribute("Email")
		EsGrupo = Usuarios.item(j).getAttribute("IsDomainGroup")
		Fichero.WriteLine (NombreGrupo+VBTab+Nombre+VBTab+Login+VBTab+Correo+VBTab+EsGrupo)
	Next 
Next 

Fichero.Close
WScript.Echo "Informe de grupos y usuarios finalizado"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub
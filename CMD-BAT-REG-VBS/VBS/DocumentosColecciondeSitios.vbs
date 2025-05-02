'VBTab = Chr(9) Tabulador

Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
        "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
        "<soap:Body>"+_
        "<GetSite xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
        "</soap:Body>"+_
        "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetSite"
URLServicio = "http://.../_vti_bin/SiteData.asmx"

EjecutaPeticion

Set Fichero = ObjetoArchivo.CreateTextFile("SalidaDocumentosColecciondeSitios.txt", True) 
Fichero.WriteLine ("Sitio"+VBTab+"IDLista"+VBTab+"Titulo"+VBTab+"Tipo"+VBTab+"Plantilla"+VBTab+"Bytes"+VBTab+"Autor"+VBTab+"Fecha Creación"+VBTab+"Documento"+VBTab+"Ruta"+VBTab+"Titulo")

'Recupera del nombre de cada uno de los Sitios
ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Sitios = ObjetoDOM.getElementsByTagName("_sWebWithTime")

For i=0 To Sitios.length-1
	Set SitioURL = Sitios.item(i).selectSingleNode("Url")
	Fichero.WriteLine (SitioURL.text)
	
	'Petición de la coleccion de listas de cada sitio
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
             "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
             "<soap:Body>"+_
			 "<GetListCollection xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
			 "</soap:Body>"+_
			 "</soap:Envelope>"
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetListCollection"
	URLServicio = SitioURL.text+"/_vti_bin/SiteData.asmx"
	
	EjecutaPeticion
	
	'Recupera datos de cada una de las listas
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set Listas = ObjetoDOM.getElementsByTagName("_sList")
	
	For j=0 to Listas.length-1
		Set ListaID = Listas.item(j).selectSingleNode("InternalName")
		Set ListaTitulo = Listas.item(j).selectSingleNode("Title")
		Set ListaTipo = Listas.item(j).selectSingleNode("BaseType")
		set ListaPlantilla = Listas.item(j).selectSingleNode("BaseTemplate")
		Fichero.WriteLine (VBTab+ListaID.text+VBTab+ListaTitulo.text+VBTab+ListaTipo.text+VBTab+ListaPlantilla.text)
		'Petición de los elementos de la lista / documentos de la biblioteca
		AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
		Peticion = "<?xml version='1.0' encoding='utf-8'?>"+_
                   "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
		           "<soap:Body>"+_
				   "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
				   "<strListName>"+ListaID.text+"</strListName>"+_
				   "</GetListItems>"+_
				   "</soap:Body>"+_
				   "</soap:Envelope>"

		EjecutaPeticion
		
		'Recupera los elementos / documento
		ObjetoDOM.loadXML(ObjetoHttp.responseText)
		Set Nodo= ObjetoDOM.selectSingleNode("//GetListItemsResult")
		ObjetoDOM.loadXML(Nodo.text)
		Set Documentos = ObjetoDOM.getElementsByTagName("z:row")
		
		For k=0 To Documentos.length-1
			Bytes = QuitarID(Documentos.item(k).getAttribute("ows_FileSizeDisplay"))
			Autor = QuitarID(Documentos.item(k).getAttribute("ows_Author"))
			FechaCreacion = Documentos.item(k).getAttribute("ows_Created")
			Nombre = Documentos.item(k).getAttribute("ows_LinkFilename")
			Ruta = QuitarID(Documentos.item(k).getAttribute("ows_FileDirRef"))
			Titulo = QuitarID(Documentos.item(k).getAttribute("ows_Title"))
			Fichero.WriteLine (VBTab+VBTab+VBTab+VBTab+VBTab+Bytes+VBTab+Autor+VBTab+FechaCreacion+VBTab+Nombre+VBTab+Ruta+VBTab+Titulo)
		Next
	Next
Next 

Fichero.Close
WScript.Echo "Informe obtenido"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Function QuitarID (Texto)
If IsNull(Texto) then
	QuitarID=""
Else
	QuitarID = Right(Texto,Len(Texto)-InStr(Texto,"#"))
End if
End Function


  
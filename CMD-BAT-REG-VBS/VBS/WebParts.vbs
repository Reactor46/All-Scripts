
Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 
Set ObjetoExplorer = CreateObject("InternetExplorer.Application")

ObjetoExplorer.navigate "about:blank" 
ObjetoExplorer.toolbar = False 
ObjetoExplorer.menubar = False 
ObjetoExplorer.visible = True 
ObjetoExplorer.statusbar = False 
ObjetoExplorer.document.write "<Body><div style='text-align:center;'><span id='Porcentaje'></span></div>"

Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
        "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
        "<soap:Body>"+_
        "<GetSite xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
        "</soap:Body>"+_
        "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetSite"
URLServicio = "http://.../_vti_bin/SiteData.asmx"

EjecutaPeticion

Set Fichero = ObjetoArchivo.CreateTextFile("SalidaWebParts.txt", True) 
Fichero.WriteLine ("Sitio"+VBTab+"Pagina"+VBTab+"Web Part"+VBTab+"Version"+VBTab+"Titulo"+VBTab+"Lista")

'Recupera del nombre de cada uno de los Sitios
ObjetoDOM.loadXML(ObjetoHttp.responseText)
Set Sitios = ObjetoDOM.getElementsByTagName("_sWebWithTime")

For i=0 To Sitios.length-1
	Porcentaje= Cint((i+1)*100/(Sitios.length)) & "% Completado"
	ObjetoExplorer.Document.GetElementById("Porcentaje").InnerHtml = Porcentaje
	Set SitioURL = Sitios.item(i).selectSingleNode("Url")
	ObjetoExplorer.document.write "<div id='Web"&i&"'>"+SitioURL.text+"</div>"
	Fichero.WriteLine (SitioURL.text)
	
	'Podrían incluise Forms GetFormCollection
	'Podrían incluise Views GetViewCollection
	
	'Paginas aspx no almacenadas en bibliotecas
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
		 	 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
		 	 "<soap:Body>"+_
			 "<EnumerateFolder xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
			 "<strFolderUrl>"+SitioURL.text+"</strFolderUrl>"+_
			 "</EnumerateFolder>"+_
			 "</soap:Body>"+_
			 "</soap:Envelope>"
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/EnumerateFolder"
	URLServicio = SitioURL.text+"/_vti_bin/SiteData.asmx"
	
	EjecutaPeticion
	
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set URLS = ObjetoDOM.getElementsByTagName("_sFPUrl")

	For j=0 to URLS.length-1
	Set URL = URLS.item(j).selectSingleNode("Url")
	If Right(URL.text,4)="aspx" Then
		AbsURL=URL.text
		Fichero.WriteLine (VBTab+AbsURL)
		ObtieneWebParts
	End If
	Next
	
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
		Porcentaje= " ... " & Cint((j+1)*100/(Listas.length)) & "% "
		ObjetoExplorer.Document.GetElementById("Web"&i).InnerHtml = SitioURL.text + Porcentaje
		Set ListaID = Listas.item(j).selectSingleNode("InternalName")
		Set ListaTipo = Listas.item(j).selectSingleNode("BaseType")
		If ListaTipo.text="DocumentLibrary" Then
			'Solamente se analizan las bibliotecas de documentos
			AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
			URLServicio = SitioURL.text+"/_vti_bin/SiteData.asmx"
			Peticion = "<?xml version='1.0' encoding='utf-8'?>"+_
					   "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
					   "<soap:Body>"+_
					   "<GetListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
					   "<strListName>"+ListaID.text+"</strListName>"+_
					   "</GetListItems>"+_
					   "</soap:Body>"+_
					   "</soap:Envelope>"

			EjecutaPeticion
			
			'Recupera los documentos
			ObjetoDOM.loadXML(ObjetoHttp.responseText)
			Set Nodo= ObjetoDOM.selectSingleNode("//GetListItemsResult")
			ObjetoDOM.loadXML(Nodo.text)
			Set Documentos = ObjetoDOM.getElementsByTagName("z:row")
			
			For k=0 To Documentos.length-1
				AbsURL = Documentos.item(k).getAttribute("ows_EncodedAbsUrl")
				AbsURL = Replace (AbsURL,SitioURL.text+"/","")
				If Right(AbsURL,4)="aspx" Then
					Fichero.WriteLine (VBTab+AbsURL)
					ObtieneWebParts
				End If				
			Next
		End If
	Next
Next 

Fichero.Close
ObjetoExplorer.Quit
WScript.Echo "Informe obtenido"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Sub ObtieneWebParts
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
		 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
		 "<soap:Body>"+_
		 "<GetWebPartProperties2 xmlns='http://microsoft.com/sharepoint/webpartpages'>"+_
		 "<pageUrl>"+AbsURL+"</pageUrl>"+_
		 "<storage>None</storage>"+_
		 "<behavior>Version3</behavior>"+_
		 "</GetWebPartProperties2>"+_
		 "</soap:Body>"+_
		 "</soap:Envelope>"
		 
AccionSOAP = "http://microsoft.com/sharepoint/webpartpages/GetWebPartProperties2"
URLServicio = SitioURL.text+"/_vti_bin/WebPartPages.asmx"

EjecutaPeticion

ObjetoDOM.loadXML(ObjetoHTTP.responseText)
Set WebParts = ObjetoDOM.getElementsByTagName("WebPart")
For l=0 To WebParts.length-1
	Set Nombre = WebParts.item(l).selectSingleNode("TypeName")
	Listatxt="":Titulotxt=""
	If Nombre is Nothing Then
		'Version 3
		Set Tipo = WebParts.item(l).selectSingleNode("//type")
		Valor=split(Tipo.getAttribute("name"),",")
		Set Propiedades = WebParts.item(l).getElementsByTagName("property")
		For m=0 to Propiedades.length-1
			If Propiedades.item(m).getAttribute("name")="ListName" Then Listatxt=Propiedades.item(m).text
			If Propiedades.item(m).getAttribute("name")="Title" Then Titulotxt=Propiedades.item(m).text
		Next
		Fichero.WriteLine (VBTab+VBTab+Valor(0)+VBTab+"v3"+VBTab+Titulotxt+VBTab+Listatxt) 
	Else
		'Version 2
		Set Titulo = WebParts.item(l).selectSingleNode("Title")
		If  not Titulo is Nothing Then Titulotxt=Titulo.text
		Set Lista = WebParts.item(l).selectSingleNode("ListName")
		If not Lista is Nothing Then Listatxt=Lista.text
		Fichero.WriteLine (VBTab+VBTab+Nombre.text+VBTab+"v2"+VBTab+Titulotxt+VBTab+Listatxt) 
	End If
Next			
End Sub




  
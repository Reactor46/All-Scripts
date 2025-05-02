
Set ObjetoStream = CreateObject("ADODB.Stream")
Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoElemento = ObjetoDOM.CreateElement("TMP")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")


Page = "Page.aspx"
LibraryDestination = "Pages/Page.aspx"
SiteURL = "http://.../"
WebPart = "Content.dwp"
HTML = "Content.html"

'Lectura de la pagina en binario
ObjetoStream.Open
ObjetoStream.type= 1 'Tipo Binario
ObjetoStream.LoadFromFile(Page)
ArchivoBinario = ObjetoStream.Read()
ObjetoStream.Close

'Conversion a Base64
ObjetoElemento.DataType = "bin.base64" 'Tipo Base64
ObjetoElemento.NodeTypedValue = ArchivoBinario
ArchivoCodificado = ObjetoElemento.Text

'Peticion de carga de la pagina
URLServicio = SiteURL+"_vti_bin/copy.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<CopyIntoItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<SourceUrl>C:/</SourceUrl>"+_
"<DestinationUrls>"+_
"<string>"+SiteURL+LibraryDestination+"</string>"+_
"</DestinationUrls>"+_
"<Fields>"+_
"<FieldInformation Type='Text' InternalName='Title' DisplayName='Titulo' Value='Archivo cargado con SOAP' />"+_
"</Fields>"+_
"<Stream>"+ArchivoCodificado+"</Stream>"+_
"</CopyIntoItems>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

'Recuperación los Ids de la pagina cargada
URLServicio = SiteURL+"_vti_bin/sitedata.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetURLSegments"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<GetURLSegments xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<strURL>"+SiteURL+LibraryDestination+"</strURL>"+_
"</GetURLSegments>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

ObjetoDOM.loadXML(ObjetoHTTP.responseText)
Set nodeBook = ObjetoDOM.selectSingleNode("//strItemID")
IDDocumento = nodeBook.text
Set nodeBook = ObjetoDOM.selectSingleNode("//strListID")
IDBiblioteca = nodeBook.text

'Actulización de la propiedad _CopySource en la pagina
URLServicio = SiteURL+"_vti_bin/lists.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<UpdateListItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<listName>"+IDBiblioteca+"</listName>"+_
"<updates>"+_
"<Batch OnError='Return'><Method ID='1' Cmd='Update'><Field Name='ID'>"+IDDocumento+"</Field><Field Name='MetaInfo' Property='_CopySource'></Field></Method></Batch>"+_
"</updates>"+_
"</UpdateListItems>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

'Lectura del archivo HTML
ObjetoDOM.load(HTML)
Set HTMLElement = ObjetoDOM.documentElement
'Lectura del archivo WebPart
ObjetoDOM.load(WebPart)
Set WebPartElement = ObjetoDOM.documentElement
'Nodo Content para agregar la seccion CDATA
Set NodeData = WebPartElement.selectSingleNode("//Content")
Set newCDATA=ObjetoDOM.createCDATASection(HTMLElement.xml)
NodeData.appendChild(newCDATA)


'Inserción  del elemento web
URLServicio = SiteURL+"_vti_bin/WebPartPages.asmx"
AccionSOAP = "http://microsoft.com/sharepoint/webpartpages/AddWebPart"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<AddWebPart xmlns='http://microsoft.com/sharepoint/webpartpages'>"+_
"<pageUrl>"+LibraryDestination+"</pageUrl>"+_
"<webPartXml>"+Encode(WebPartElement.xml)+"</webPartXml>"+_
"<storage>Shared</storage>"+_
"</AddWebPart>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

Wscript.Echo "Página con elemento web publicada"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Function Encode (Texto)
Encode=Replace(Texto,"<","&lt;")
Encode=Replace(Encode,">","&gt;")
Encode=Replace(Encode,chr(34),"&quot;")
End Function


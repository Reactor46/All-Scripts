Const ArchivoLocal = "Ejemplo.pptx"
Const URLDestino = "http://.../Biblioteca/Ejemplo.pptx"

Set ObjetoStream = CreateObject("ADODB.Stream")
Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoElemento = ObjetoDOM.CreateElement("TMP")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")


'Lectura del archivo en binario
ObjetoStream.Open
ObjetoStream.type= 1 'Tipo Binario
ObjetoStream.LoadFromFile(ArchivoLocal)
ArchivoBinario = ObjetoStream.Read()
ObjetoStream.Close

'Conversion a Base64
ObjetoElemento.DataType = "bin.base64" 'Tipo Base64
ObjetoElemento.NodeTypedValue = ArchivoBinario
ArchivoCodificado = ObjetoElemento.Text

'Construye texto Peticion de carga del documento
URLServicio = "http://.../_vti_bin/copy.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<CopyIntoItems xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<SourceUrl>C:/</SourceUrl>"+_
"<DestinationUrls>"+_
"<string>"+URLDestino+"</string>"+_
"</DestinationUrls>"+_
"<Fields>"+_
"<FieldInformation Type='Text' InternalName='Title' DisplayName='Titulo' Value='Archivo cargado con SOAP' />"+_
"</Fields>"+_
"<Stream>"+ArchivoCodificado+"</Stream>"+_
"</CopyIntoItems>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion


'Peticion para recuperar los Ids del documento cargado
URLServicio = "http://.../_vti_bin/sitedata.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetURLSegments"
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<GetURLSegments xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<strURL>"+URLDestino+"</strURL>"+_
"</GetURLSegments>"+_
"</soap:Body>"+_
"</soap:Envelope>"

EjecutaPeticion

ObjetoDOM.loadXML(ObjetoHTTP.responseText)
Set nodeBook = ObjetoDOM.selectSingleNode("//strItemID")
IDDocumento = nodeBook.text
Set nodeBook = ObjetoDOM.selectSingleNode("//strListID")
IDBiblioteca = nodeBook.text


'Peticion para actualizar _CopySource en el documento
URLServicio = "http://.../_vti_bin/lists.asmx"
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

WScript.Echo "Documento publicado"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub


Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHttp = CreateObject("Microsoft.XMLHTTP")
Set ObjetoStream = CreateObject("ADODB.Stream")

'Construye texto Peticion de descarga del documento
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
"<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
"<soap:Body>"+_
"<GetItem xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"+_
"<Url>http://../Biblioteca/Ejemplo.pptx</Url>"+_
"</GetItem>"+_
"</soap:Body>"+_
"</soap:Envelope>"

URLServicio = "http://../_vti_bin/Copy.asmx"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetItem"

ObjetoHttp.open "Get", URLServicio, false
ObjetoHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHttp.setRequestHeader "SOAPAction", AccionSOAP
ObjetoHttp.send Peticion

'Recupera el la respuesta
ObjetoDOM.loadXML(ObjetoHttp.responseText)

'Obtiene el valor del nodo Stream y lo convierte a binario
Set nodeBook = ObjetoDOM.selectSingleNode("//Stream")
nodeBook.DataType = "bin.base64" 'Tipo Base64
ArchivoBinario = nodeBook.NodeTypedValue

'Guarda el archivo localmente
ObjetoStream.Open
ObjetoStream.Type = 1 'Tipo Binario
ObjetoStream.Write (ArchivoBinario)
ObjetoStream.SaveToFile("Ejemplo.pptx")




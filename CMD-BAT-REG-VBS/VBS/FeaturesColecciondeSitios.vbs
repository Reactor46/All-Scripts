
Set ObjetoDOM = CreateObject("Microsoft.XMLDOM")
Set ObjetoHTTP = CreateObject("Microsoft.XMLHTTP")
Set ObjetoArchivo = CreateObject("Scripting.FileSystemObject") 

URLColeccion="http://..."

'Recupera el GUID y Propiedades del archivo Features.txt
Set Archivo = ObjetoArchivo.OpenTextFile("Features.txt", 1)
Dim Features(250,1)
i = 0
Do Until Archivo.AtEndOfStream
tmp=split(Archivo.ReadLine,VBTab)
Features(i,0)=tmp(0)
Features(i,1)=tmp(1)+VBTab+tmp(2)+VBTab+tmp(3)
i = i + 1
Loop
Archivo.Close

'Recorre todos los sitios y subsitios de la colección de sitios
Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
         "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
         "<soap:Body>"+_
         "<GetSite xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
         "</soap:Body>"+_
         "</soap:Envelope>"
AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetSite"
URLServicio = URLColeccion+"/_vti_bin/SiteData.asmx"

EjecutaPeticion

Set Archivo = ObjetoArchivo.CreateTextFile("SalidaFeatureSiteCollection.txt", True) 
Archivo.WriteLine ("Sitio"+VBTab+"Features Activas"+VBTab+"Nombre"+VBTab+"Título"+VBTab+"Ámbito")
'Recupera el nombre cada subsitio
ObjetoDOM.loadXML(ObjetoHttp.responseText)
WScript.Echo ObjetoHttp.responseText
Set Sitios = ObjetoDOM.getElementsByTagName("_sWebWithTime")

For i=0 To Sitios.length-1
	Set SitioURL = Sitios.item(i).selectSingleNode("Url")
	Archivo.WriteLine (SitioURL.text)
	'Petición de las Features activas del sitio
	Peticion="<?xml version='1.0' encoding='utf-8'?>"+_
			 "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"+_
			 "<soap:Body>"+_
			 "<GetActivatedFeatures xmlns='http://schemas.microsoft.com/sharepoint/soap/' />"+_
			 "</soap:Body>"+_
			 "</soap:Envelope>"
	AccionSOAP = "http://schemas.microsoft.com/sharepoint/soap/GetActivatedFeatures"
	URLServicio = SitioURL.text+"/_vti_bin/Webs.asmx"	
	EjecutaPeticion
	ObjetoDOM.loadXML(ObjetoHttp.responseText)
	Set FeaturesNodo=ObjetoDOM.selectSingleNode("//GetActivatedFeaturesResult")
	If i=0 Then
		'URL de la Colección: Ambito: Site
		tmp = Split(Replace(FeaturesNodo.text,VBTab,""),",")
		BuscaPropiedades(tmp)
	Else
		'URL de los Sitios: Ambito: Web
		tmp = Split(FeaturesNodo.text,VBTab)
		tmpsite = Split(tmp(0),",")
		BuscaPropiedades(tmpsite)
	End If
Next
Archivo.Close

WScript.Echo "Informe obtenido"

Private Sub EjecutaPeticion
ObjetoHTTP.Open "Get", URLServicio, false
ObjetoHTTP.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
ObjetoHTTP.SetRequestHeader "SOAPAction", AccionSOAP
ObjetoHTTP.Send Peticion
End Sub

Private Sub BuscaPropiedades (Array)
	For Each FeatureID in Array
			Propiedades=""
			For j=0 To UBound(Features)			
			If Features(j,0)=FeatureID Then
			Propiedades=Features(j,1)
			Exit For
			End If
			Next
			Archivo.WriteLine (VBTab+FeatureID+VBTab+Propiedades)
	Next
End Sub

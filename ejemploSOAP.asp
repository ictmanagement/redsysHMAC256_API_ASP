<%@ language="vbscript" codepage="65001"%>
<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>


<!-- #include file="./include/dvim_apiRedsys_VB.asp" -->

<% 
	Response.CharSet = "utf-8"
	Response.ContentType = "text/xml"

	If (Request("wsdl").Count>0) Then
		Server.Transfer "InotificacionSIS.wsdl"
		Response.End
	End If

	Dim reqSize, content
	reqSize=Request.TotalBytes
	content = Request.BinaryRead(reqSize)

	Dim sSOAPOut
	sSOAPOut = procesaNotificacionSIS(content)
	Response.Write sSOAPOut

	' notificacionSIS_request es Byte() con el NotificacionSIS SOAP Request
	' devuelve strNotificacionSIS_response (utf8)String'
	Function procesaNotificacionSIS(ByVal notificacionSIS)
		' Se crea Objeto
		Dim miObj
		Set miObj = new RedsysAPI
		' Obtenemos el String UTF8 del notificacionSIS
		Dim strInput
		strInput = miObj.ConvertToUtf8String(notificacionSIS)
		If (Len(strInput)>0) Then
			' Extraemos el SOAP envelope
			Dim xmlMessage
			xmlMessage = extractSOAPEnvelope(strInput)
			Dim signatureRecibida
			signatureRecibida = miObj.getSignatureNotifSOAP(xmlMessage)
			Dim kc, firma
			kc = "Mk9m98IfEblmPfrpsawt7BmxObt98Jev" 'Clave recuperada de CANALES
			firma = miObj.createMerchantSignatureNotifSOAPRequest(kc,xmlMessage) 'TODO: encodedMessage o de xmlMessage ?'
			Dim res
			res = "KO"
			If (firma = signatureRecibida) Then 
				'Aquí deberíais verificar los parámetros recibidos, pe 
				'Ds_Card_Country -> miObj.getParameter("Card_Country"), 
				'Ds_Response -> miObj.getParameter("Response"), 
				'Ds_AuthorisationCode -> miObj.getParameter("AuthorisationCode"), etc.
				res = "OK"
			End If
			Dim numPedido
			numPedido = miObj.getParameter("Order") 'Ds_Order
			Dim xmlResp 
			xmlResp = "<Response Ds_Version='0.0'><Ds_Response_Merchant>" & res & "</Ds_Response_Merchant></Response>"
			Dim xmlSign
			xmlSign = "<Signature>" & miObj.createMerchantSignatureNotifSOAPResponse(kc,xmlResp,numPedido) & "</Signature>"
			Dim xmlMessageResponse
			xmlMessageResponse = "<Message>" & xmlResp & xmlSign & "</Message>"
			procesaNotificacionSIS=generateSOAPEnvelope(xmlMessageResponse)
		Else
			'TODO: generar el SOAP error correspondiente
			procesaNotificacionSIS=generateSOAPEnvelope("")
		End If
	End Function

	'Función básica para generar el envoltorio SOAP al mensaje de respuesta
	Function generateSOAPEnvelope(xmlMessageResponse)
		Dim encodedMessage
		encodedMessage = XML1_encode (xmlMessageResponse)
		generateSOAPEnvelope = "" & _
"<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inot=""InotificacionSIS"">" & _
"<soapenv:Header/>" & _
"<soapenv:Body>" & _
"<inot:procesaNotificacionSISResponse soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" & _
"<return xsi:type=""xsd:string"">" & encodedMessage & "</return>" & _
"</inot:procesaNotificacionSISResponse>" & _
"</soapenv:Body>" & _
"</soapenv:Envelope>"
	End Function

		'Función básica para eliminar el envoltorio SOAP al mensaje de solicitud
	Function extractSOAPEnvelope(strSOAPRequest)
		Dim posXMLIni, posXMLFin, posXMLIniEnd, encodedXMLContent
		posXMLIni = InStr(strSOAPRequest, "<XML")
		posXMLFin = InStr(strSOAPRequest, "</XML>")
		encodedXMLContent = Mid(strSOAPRequest,posXMLIni,posXMLFin-posXMLIni)
		posXMLIniEnd = InStr(encodedXMLContent, ">")
		encodedXMLContent = Mid(encodedXMLContent, posXMLIniEnd+1)
		extractSOAPEnvelope = XML1_decode(encodedXMLContent)
	End Function
	
	Function XML1_decode(encodedXML)
		Dim decodedXML
		decodedXML=encodedXML
		'Tabla corresponsdiente a PHP get_html_translation_table(HTML_ENTITIES, ENT_QUOTES | ENT_XML1)
		decodedXML=Replace(decodedXML,"&quot;","""")
		decodedXML=Replace(decodedXML,"&apos;","'")
		decodedXML=Replace(decodedXML,"&lt;","<")
		decodedXML=Replace(decodedXML,"&gt;",">")
		decodedXML=Replace(decodedXML,"&amp;","&")
		XML1_decode = decodedXML
	End Function
	Function XML1_encode(decodedXML)
		Dim encodedXML
		encodedXML=decodedXML
		'Tabla corresponsdiente a PHP get_html_translation_table(HTML_ENTITIES, ENT_QUOTES | ENT_XML1)
		encodedXML=Replace(encodedXML,"&","&amp;")
		encodedXML=Replace(encodedXML,"""","&quot;")
		encodedXML=Replace(encodedXML,"'","&apos;")
		encodedXML=Replace(encodedXML,"<","&lt;")
		encodedXML=Replace(encodedXML,">","&gt;")
		XML1_encode = encodedXML
	End Function
	
%>

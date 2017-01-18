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
	call procesaNotificacionSIS(content,sOut)
	Response.Write generateSOAPEnvelope(sOut)

	Sub procesaNotificacionSIS(ByVal sInput, ByRef sOutput)
		If (Len(sInput)>0) Then'//URL DE RESP. ONLINE
			' Se crea Objeto
			Set miObj = new RedsysAPI
			Dim signatureRecibida
			signatureRecibida = miObj.getSignatureNotifSOAP(sInput)
			Dim kc, firma
			kc = "Mk9m98IfEblmPfrpsawt7BmxObt98Jev" 'Clave recuperada de CANALES
			firma = miObj.createMerchantSignatureNotifSOAPRequest(kc,sInput)
			Dim res
			If (firma = signatureRecibida) Then 
				'Aquí deberíais verificar los parámetros recibidos, pe miObj.getParameter("Ds_Card_Country"), miObj.getParameter("Ds_Response"), etc
				res = "OK"
			Else
				res = "KO"
			End If
			Dim numPedido
			numPedido = miObj.getParameter("Ds_Order")
			Dim resp 
			resp = "<Response Ds_Version=""0.0""><Ds_Response_Merchant>" & res & "</Ds_Response_Merchant></Response>"
			Dim sign
			sign = miObj.createMerchantSignatureNotifSOAPResponse(kc,resp,numPedido)
			sign = "<Signature>" & sign & "</Signature>"
			sOutput = "<Message>" & resp & sign & "</Message>"
		End If
	End Sub

	'Función básica para generar el envoltorio SOAP al mensaje de respuesta
	Function generateSOAPEnvelope(strReturn)
		generateSOAPEnvelope = "" & _
"<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:inot=""InotificacionSIS"">" & _
"<soapenv:Header/>" & _
"<soapenv:Body>" & _
"<inot:procesaNotificacionSISResponse soapenv:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" & _
"<return xsi:type=""xsd:string"">" & strReturn & "</return>" & _
"</inot:procesaNotificacionSISResponse>" & _
"</soapenv:Body>" & _
"</soapenv:Envelope>"
	End Function

%>

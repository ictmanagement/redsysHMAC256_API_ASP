<%@ language="vbscript" codepage="65001"%>
<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>


<!-- #include file="./include/dvim_apiRedsys_VB.asp" -->

<html lang="es">
<head>
	<title>Ejemplo de Recepción de Respuesta a Petición (URL OK, KO)</title>
</head>
<body>
<%
	Response.CharSet = "utf-8"

	' Se crea Objeto
	Dim miObj 
	Set miObj = new RedsysAPI

	If (Request.Form.Count>0 or Request.QueryString.Count>0) Then'//URL DE RESP. ONLINE
		Dim version, datos, signatureRecibida					
		version = Request("Ds_SignatureVersion")
		datos = Request("Ds_MerchantParameters")
		signatureRecibida = Request("Ds_Signature")

		Dim kc, firma
		kc = "Mk9m98IfEblmPfrpsawt7BmxObt98Jev" 'Clave recuperada de CANALES
		firma = miObj.createMerchantSignatureNotif(kc,datos)

		If (firma = signatureRecibida) Then
			Response.Write "FIRMA OK<br/>"
		Else
			Response.Write "FIRMA KO<br/>"
		End If
		Response.Write "Ds_Date=" & miObj.getParameter("Ds_Date") & "<br/>"
		Response.Write "Ds_Hour=" & miObj.getParameter("Ds_Hour") & "<br/>"
		Response.Write "Ds_Order=" & miObj.getParameter("Ds_Order") & "<br/>"
		Response.Write "Ds_SecurePayment=" & miObj.getParameter("Ds_SecurePayment") & "<br/>"
		Response.Write "Ds_Amount=" & miObj.getParameter("Ds_Amount") & "<br/>"
		Response.Write "Ds_Currency=" & miObj.getParameter("Ds_Currency") & "<br/>"
		Response.Write "Ds_MerchantCode=" & miObj.getParameter("Ds_MerchantCode") & "<br/>"
		Response.Write "Ds_Terminal=" & miObj.getParameter("Ds_Terminal") & "<br/>"
		Response.Write "Ds_Response=" & miObj.getParameter("Ds_Response") & "<br/>"
		Response.Write "Ds_TransactionType=" & miObj.getParameter("Ds_TransactionType") & "<br/>"
		Response.Write "Ds_MerchantData=" & miObj.getParameter("Ds_MerchantData") & "<br/>"
		Response.Write "Ds_AuthorisationCode=" & miObj.getParameter("Ds_AuthorisationCode") & "<br/>"
		Response.Write "Ds_ConsumerLanguage=" & miObj.getParameter("Ds_ConsumerLanguage") & "<br/>"

	End If

%>
</body> 
</html> 
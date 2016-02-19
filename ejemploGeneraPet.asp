<%@ language="vbscript" codepage="65001"%>
<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>


<!-- #include file="./include/dvim_apiRedsys_VB.asp" -->

<html lang="es">
<head>
	<title>Ejemplo de Generación de Petición</title>
</head>
<body>
<% 
	Response.CharSet = "utf-8"

	' Se crea Objeto
	Dim miObj 
	Set miObj = new RedsysAPI
		
	' Valores de entrada
	Dim fuc,terminal,moneda,trans,url,urlOKKO,id,amount
	fuc="999008881"
	terminal="871"
	moneda="978"
	trans="0"
	amount="145"
	id="ped" & time() 'no acepta +,&,€,"
	url="" 'Colocar la URL completa de ejemploSOAP.asp (debe ser accesible desde RedSys, por tanto no puede ser localhost)
	urlOKKO="" 'Colocar la URL completa de ejemploRecepcionaPet.asp (puede ser localhost para pruebas)
	
	' Se Rellenan los campos
	call miObj.setParameter("DS_MERCHANT_AMOUNT",amount)
	call miObj.setParameter("DS_MERCHANT_ORDER",CStr(id))
	call miObj.setParameter("DS_MERCHANT_MERCHANTCODE",fuc)
	call miObj.setParameter("DS_MERCHANT_CURRENCY",moneda)
	call miObj.setParameter("DS_MERCHANT_TRANSACTIONTYPE",trans)
	call miObj.setParameter("DS_MERCHANT_TERMINAL",terminal)
	call miObj.setParameter("DS_MERCHANT_MERCHANTURL",url)
	call miObj.setParameter("DS_MERCHANT_URLOK",urlOKKO)	
	call miObj.setParameter("DS_MERCHANT_URLKO",urlOKKO)

	' Datos de configuración
	Dim version
	version="HMAC_SHA256_V1"
	kc = "Mk9m98IfEblmPfrpsawt7BmxObt98Jev" 'Clave recuperada de CANALES
	' Se generan los parámetros de la petición
	Dim request,params,signature
	request = ""
	params = miObj.createMerchantParameters()
	signature = miObj.createMerchantSignature(kc)

	Dim postURL
	postURL = "https://sis-d.redsys.es/sis/realizarPago"  'URL DE DESARROLLO CON HTTPS
	'postURL = "http://sis-d.redsys.es/sis/realizarPago" 'URL DE DESARROLLO
	'postURL = "https://sis-t.redsys.es:25443/sis/realizarPago"  'URL DE PRUEBAS, USAR CON LOS DATOS DE VUESTRO COMERCIO
	'postURL = "https://sis.redsys.es/sis/realizarPago"  'URL DE PRODUCCION, USAR CON LOS DATOS DE VUESTRO COMERCIO

	 
%>
<form name="frm" action="<%=postURL%>" method="POST" target="_blank">
Ds_Merchant_SignatureVersion <input type="text" name="Ds_SignatureVersion" value="<%=version%>"/><br/>
Ds_Merchant_MerchantParameters <input type="text" name="Ds_MerchantParameters" value="<%=params%>"/><br/>
Ds_Merchant_Signature <input type="text" name="Ds_Signature" value="<%=signature%>"/><br/>
<input type="submit" value="Enviar" >
</form>

</body>
</html>

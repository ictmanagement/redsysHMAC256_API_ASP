<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>

<!-- #include file="./dvim_brix_crypto-js-master_VB.asp" -->

<!-- #include file="./dvim_json_douglascrockford_JSON-js_JS.asp" -->

<script language="vbscript" runat="server">

	Class RedsysAPI

		'/******  Array de DatosEntrada ******/
		Private vars_pay
		Private  objJSON

		Private Sub Initialize()
			Set vars_pay = Nothing
			Set objJSON = Nothing
			Set objJSON = JSON
			Set vars_pay = CreateObject("Scripting.Dictionary")
			vars_pay.CompareMode=1  'TextCompare
		End Sub

		Private Sub Class_Initialize()
			Initialize
		End Sub
		Private Sub Class_Terminate()
			Set vars_pay = Nothing
			Set objJSON = Nothing
		End Sub
		'/******  Set parameter ******/
		Public Sub SetParameter(key, value)
			vars_pay.Item(ucase(key)) = value
		End Sub
		'/******  Get parameter ******/
		Public Function GetParameter(key)
			If vars_pay.Exists(key) Then
				GetParameter = vars_pay.Item(key)
			Else
				GetParameter = Empty
			End if
		End Function

		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'////////////					FUNCIONES AUXILIARES:							  ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'recibe (Utf8)String|WordArray y retorna WordArray
		Private Function ConvertUtf8StrToWordArray(data)
			If (typename(data) = "String") Then
				Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data)
			Elseif (typename(data) = "JScriptTypeInfo") Then
				On error resume next
				'Set ConvertUtf8StrToWordArray = data.clone() 'solo verifica que tenga clone
				'Set ConvertUtf8StrToWordArray = CryptoJS.enc.Utf8.parse(data.toString(CryptoJS.enc.Utf8)) 'verifica que sea WordArray y que admita utf8
				Set ConvertUtf8StrToWordArray = CryptoJS.lib.WordArray.create().concat(data) 'verifica que sea WordArray
				If Err.number>0 Then
					Set ConvertUtf8StrToWordArray = Nothing
				End if
				On error goto 0
			Else
				Set ConvertUtf8StrToWordArray = Nothing
			End if
		End Function
		'recibe (Utf8)String|WordArray|Byte() y retorna (Utf8)String
		Public Function ConvertToUtf8String(data)
			If TypeName(data) = "Byte()" Then
				Set data = CryptoJS.enc.Hex.parse(ByteArrayToHexString(data))
			End If
			If (typename(data) = "String") Then
				ConvertToUtf8String = CryptoJS.enc.Utf8.parse(data).toString(CryptoJS.enc.Utf8)
			Elseif (typename(data) = "JScriptTypeInfo") Then
				On error resume next
				ConvertToUtf8String = data.toString(CryptoJS.enc.Utf8)
				If Err.number>0 Then
					ConvertToUtf8String = ""
				End if
				On error goto 0
			Else
				ConvertToUtf8String = ""
			End if
		End Function
		'/******  3DES Function  ******/
		'recibe (Utf8)String|WordArray y retorna WordArray (CipherParams.ciphertext)
		Private Function encrypt_3DES(message, key) 
			'Se establece un IV por defecto
			Dim iv 
			Set iv = CryptoJS.enc.Hex.parse("0000000000000000")
			Dim messageWA
			Set messageWA = ConvertUtf8StrToWordArray(message)
			Dim keyWA
			Set keyWA = ConvertUtf8StrToWordArray(key)
			' Se cifra
			Dim cfg 
			Set cfg = new Encrypt_Cfg  
			Set cfg.iv=iv
			Set cfg.mode=CryptoJS.mode.CBC
			Set cfg.padding=CryptoJS.pad.ZeroPadding 
			Dim ciphertext
			Set ciphertext = CryptoJS.TripleDES.encrypt(messageWA, keyWA, cfg).ciphertext
			Set encrypt_3DES = ciphertext
		End Function
		'/******  Base64 Functions  ******/
		' recibe (Utf8)String|WordArray y retorna (base64)String
		Private Function base64_url_encode(input)
			base64_url_encode = Replace(Replace(encodeBase64(input),"+","-"),"/","_")
		End Function
		' recibe (Utf8)String|WordArray y retorna (base64)String
		Private Function encodeBase64(data) 
			Dim dataWA
			Set dataWA = ConvertUtf8StrToWordArray(data)
			Dim encodedData
			encodedData = CryptoJS.enc.Base64.stringify(dataWA)
			encodeBase64 = encodedData
		End Function
		' recibe (base64)String|WordArray y retorna WordArray
		Private Function base64_url_decode(input)
			Dim inputStr
			inputStr = ConvertToUtf8String(input)
			Set base64_url_decode = decodeBase64(Replace(Replace(inputStr,"-","+"),"_","/"))
		End Function
		' recibe (base64)String|WordArray y retorna WordArray
		Private Function decodeBase64(data)
			Dim decodedDataWA
			Set decodedDataWA = CryptoJS.enc.Base64.parse(data)
			Set decodeBase64 = decodedDataWA
		End Function
		'/******  MAC Function ******/
		'recibe String|WordArray , retorna WordArray
		Private Function mac256(ent, key) 
			Dim encWA
			Set encWA = ConvertUtf8StrToWordArray(ent)
			Dim keyWA
			Set keyWA = ConvertUtf8StrToWordArray(key)
			Dim resWA
			Set resWA = CryptoJS.HmacSHA256(encWA, keyWA)
			Set mac256 = resWA
		End Function

		'//////////////////////////////////////////////////////////////////////////////////////////////
		'////////////	   FUNCIONES PARA LA GENERACIÓN DEL FORMULARIO DE PAGO:			  ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////

		'//////////////////////////////////////////////////////////////////////////////////////////////
		'/////////' FUNCIONES PARA LA RECEPCIÓN DE DATOS DE PAGO (Notif, URLOK y URLKO): ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'/******  Obtener Número de pedido ******/
		Private Function getOrder()
			Dim numPedido 
			numPedido = GetParameter("DS_MERCHANT_ORDER")
			If IsEmpty(numPedido) Then
				numPedido = ""
			End If 
			getOrder = numPedido
		End Function
		'/******  Convertir Dictionary en (JSON)String ******/
		Private Function dictionatyToJsonString()
			dictionatyToJsonString = objJSON.stringify(vars_pay,objJSON.replacer)
		End Function
		Public Function createMerchantParameters()
			' Se transforma el Dictionary en un string Json
			Dim json
			json = dictionatyToJsonString()
			' Se codifican los datos Base64
			createMerchantParameters = encodeBase64(json)
		End Function
		Public Function createMerchantSignature(key) 
			'Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			'Se genera el parámetro Ds_MerchantParameters
			Dim ent
			ent = createMerchantParameters()
			'Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(getOrder(), keyWA)
			'MAC256 del parámetro Ds_MerchantParameters
			Dim resWA
			Set resWA = mac256(ent, keyWA)
			'Se codifican los datos Base64
			createMerchantSignature= encodeBase64(resWA)
		End Function
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////// FUNCIONES PARA LA RECEPCIÓN DE DATOS DE PAGO (Notif, URLOK y URLKO): ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'/******  Obtener Número de pedido ******/
		Private Function getOrderNotif()
			Dim numPedido 
			numPedido = GetParameter("DS_ORDER")
			If IsEmpty(numPedido) Then
				numPedido = ""
			End If 
			getOrderNotif = numPedido
		End Function
		Private Function getOrderNotifSOAP(datos)
			Dim posPedidoIni, tamPedidoIni, posPedidoFin
			posPedidoIni = InStr(datos, "<Ds_Order>")
			tamPedidoIni = Len("<Ds_Order>")
			posPedidoFin = InStr(datos, "</Ds_Order>")
			getOrderNotifSOAP = Mid(datos,posPedidoIni + tamPedidoIni,posPedidoFin - (posPedidoIni + tamPedidoIni))
		End Function
		Private Function getRequestNotifSOAP(datos)
			Dim posReqIni, tamReqFin, posReqFin
			posReqIni = InStr(datos, "<Request")
			posReqFin = InStr(datos, "</Request>")
			tamReqFin = Len("</Request>")
			getRequestNotifSOAP = Mid(datos,posReqIni,(posReqFin + tamReqFin) - posReqIni)
		End Function
		Private Function getResponseNotifSOAP(datos)
			Dim posReqIni, tamReqFin, posReqFin
			posReqIni = InStr(datos, "<Response")
			posReqFin = InStr(datos, "</Response>")
			tamReqFin = Len("</Response>")
			getResponseNotifSOAP = Mid(datos,posReqIni,(posReqFin + tamReqFin) - posReqIni)
		End Function
		Public Function getSignatureNotifSOAP(datos)
			Dim datosStr
			datosStr = ConvertToUtf8String(datos)
			Dim posSignatureIni, tamSignatureIni, posSignatureFin
			posSignatureIni = InStr(datosStr, "<Signature>")
			tamSignatureIni = Len("<Signature>")
			posSignatureFin = InStr(datosStr, "</Signature>")
			getSignatureNotifSOAP = Mid(datosStr,posSignatureIni + tamSignatureIni,posSignatureFin - (posSignatureIni + tamSignatureIni))
		End Function
		'/******  Cargar (JSON)String al Dictionary ******/
		Private Sub loadJsonStringToDictionary(datosDecod)
			Initialize
			Set vars_pay = objJSON.parse(datosDecod,objJSON.reviver)
		End Sub
		Private Function decodeMerchantParameters(datos)
			' Se decodifican los datos Base64
			decodeMerchantParameters = base64_url_decode(datos).toString(CryptoJS.enc.Utf8)
		End Function
		Public Function createMerchantSignatureNotif(key, datos)
			'Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			' Se decodifican los datos Base64
			Dim decodec
			decodec = decodeMerchantParameters(datos)
			' Los datos decodificados se pasan al Dictionaty
			loadJsonStringToDictionary(decodec)
			'Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(getOrderNotif(), keyWA)
			' MAC256 del parámetro Ds_Parameters que envía Redsys
			Dim resWA
			Set resWA = mac256(datos, keyWA)
			'Se codifican los datos Base64
			createMerchantSignatureNotif = base64_url_encode(resWA)
		End Function

		'/******  Notificaciones SOAP ENTRADA ******/
		Public Function createMerchantSignatureNotifSOAPRequest(key, datos)
			' Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			Dim datosStr
			datosStr = ConvertToUtf8String(datos)
			' Se obtienen los datos del Request
			Dim datosReq
			datosReq = getRequestNotifSOAP(datosStr)
			' Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(getOrderNotifSOAP(datosReq), keyWA)
			' MAC256 del parámetro Ds_Parameters que envía Redsys
			Dim resWA
			Set resWA = mac256(datosReq, keyWA)
			' Se codifican los datos Base64
			createMerchantSignatureNotifSOAPRequest = encodeBase64(resWA)
			' Los datos del Request se pasan al Dictionaty
			loadXMLStringToDictionary(datosReq)
		End Function
		'/******  Notificaciones SOAP SALIDA ******/
		Public Function createMerchantSignatureNotifSOAPResponse(key, datos, numPedido)
			' Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			Dim datosStr
			datosStr = ConvertToUtf8String(datos)
			' Se obtienen los datos del Response
			Dim datosRes
			datosRes = getResponseNotifSOAP(datosStr)
			' Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(numPedido, keyWA)
			' MAC256 del parámetro Ds_Parameters que envía Redsys
			Dim resWA
			Set resWA = mac256(datosRes, keyWA)
			' Se codifican los datos Base64
			createMerchantSignatureNotifSOAPResponse = encodeBase64(resWA)
		End Function
		'/******  Cargar (XML)String al Dictionary ******/
		Private Sub loadXMLStringToDictionary(data)
			Initialize
			GetParameterDiccionary data,"Fecha","<Fecha>","</Fecha>"
			GetParameterDiccionary data,"Hora","<Hora>","</Hora>"
			GetParameterDiccionary data,"SecurePayment","<Ds_SecurePayment>","</Ds_SecurePayment>"
			GetParameterDiccionary data,"Card_Country","<Ds_Card_Country>","</Ds_Card_Country>"
			GetParameterDiccionary data,"Amount","<Ds_Amount>","</Ds_Amount>"
			GetParameterDiccionary data,"Currency","<Ds_Currency>","</Ds_Currency>"
			GetParameterDiccionary data,"Order","<Ds_Order>","</Ds_Order>"
			GetParameterDiccionary data,"MerchantCode","<Ds_MerchantCode>","</Ds_MerchantCode>"
			GetParameterDiccionary data,"Terminal","<Ds_Terminal>","</Ds_Terminal>"
			GetParameterDiccionary data,"Response","<Ds_Response>","</Ds_Response>" 
			GetParameterDiccionary data,"MerchantData","<Ds_MerchantData>","</Ds_MerchantData>"
			GetParameterDiccionary data,"TransactionType","<Ds_TransactionType>","</Ds_TransactionType>"
			GetParameterDiccionary data,"ConsumerLanguage","<Ds_ConsumerLanguage>","</Ds_ConsumerLanguage>"
			GetParameterDiccionary data,"AuthorisationCode","<Ds_AuthorisationCode>","</Ds_AuthorisationCode>"
		End Sub            
		Private Sub GetParameterDiccionary(data, key, keyInit, keyEnd)
			Dim posIni, tamIni, posFin
			posIni = InStr(data, keyInit)
			tamIni = Len(keyInit)
			posFin = InStr(data, keyEnd)
			If ( posIni > 0 and posFin >0) Then
				Dim res
				res = Mid(data,posIni + tamIni, posFin - (posIni + tamIni))
				SetParameter key, res
			End If
		End Sub    
		Private Function ByteArrayToHexString(data)
			Dim i, item, strLine, hexLine
			hexLine = ""
			For i = 0 To LenB(data)-1
				item = AscB(MidB(data,i+1,1))
				hexLine = hexLine & right("0" & hex(item),2)
			Next
			ByteArrayToHexString = hexLine
		End Function        

	End Class

</script>
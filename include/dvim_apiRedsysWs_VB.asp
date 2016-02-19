<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>

<!-- #include file="./dvim_brix_crypto-js-master_VB.asp" -->

<!-- #include file="./dvim_json_douglascrockford_JSON-js_JS.asp" -->

<script language="vbscript" runat="server">

	Class RedsysAPIWs

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
		'recibe (Utf8)String|WordArray y retorna (Utf8)String
		Private Function ConvertToUtf8String(data)
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
		'////////////	   FUNCIONES PARA LA GENERACIÓN LA PETICIÓN DE PAGO:			  ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////

		'//////////////////////////////////////////////////////////////////////////////////////////////
		'/////////' FUNCIONES PARA LA RECEPCIÓN DE DATOS DE PAGO (Notif, URLOK y URLKO): ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'/******  Obtener Número de pedido ******/
		Private Function getOrder(datos)
			getOrder = ""
			Dim posPedidoIni, tamPedidoIni, posPedidoFin
			posPedidoIni = InStr(datos, "<DS_MERCHANT_ORDER>")
			tamPedidoIni = Len("<DS_MERCHANT_ORDER>")
			posPedidoFin = InStr(datos, "</DS_MERCHANT_ORDER>")
			numPedido = Mid(datos,posPedidoIni + tamPedidoIni,posPedidoFin - (posPedidoIni + tamPedidoIni))
			getOrder = numPedido
		End Function
		Function createMerchantSignatureHostToHost(key,ent) 
			'Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			'Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(getOrder(ent), keyWA)
			'MAC256 del parámetro Ds_MerchantParameters
			Dim resWA
			Set resWA = mac256(ent, keyWA)
			'Se codifican los datos Base64
			createMerchantSignature= encodeBase64(resWA)
		End Function
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////// FUNCIONES PARA LA RECEPCIÓN DE DATOS DE PAGO (Respuesta HOST to HOST) ////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		'//////////////////////////////////////////////////////////////////////////////////////////////
		Function createMerchantSignatureResponseHostToHost(key, datos, numpedido)
			' Se decodifica la clave Base64
			Dim keyWA
			Set keyWA = decodeBase64(key)
			' Se diversifica la clave con el Número de Pedido
			Set keyWA = encrypt_3DES(numPedido, keyWA)
			' MAC256 del parámetro Ds_Parameters que envía Redsys
			Dim resWA
			Set resWA = mac256(datos, keyWA)
			' Se codifican los datos Base64
			createMerchantSignatureNotifSOAPResponse = encodeBase64(resWA)
		End Function
		'/******  Cargar (XML)String al Dictionary ******/
		Public Sub XMLToDiccionary(data)
			Class_Initialize
			GetParameterDiccionary data,"CODIGO","<CODIGO>","</CODIGO>"
			GetParameterDiccionary data,"Ds_Amount","<Ds_Amount>","</Ds_Amount>"
			GetParameterDiccionary data,"Ds_Currency","<Ds_Currency>","</Ds_Currency>"
			GetParameterDiccionary data,"Ds_Order","<Ds_Order>", "</Ds_Order>"
			GetParameterDiccionary data,"Ds_Signature","<Ds_Signature>","</Ds_Signature>"
			GetParameterDiccionary data,"Ds_MerchantCode","<Ds_MerchantCode>", "</Ds_MerchantCode>"
			GetParameterDiccionary data,"Ds_Terminal","<Ds_Terminal>", "</Ds_Terminal>"
			GetParameterDiccionary data, "Ds_Response", "<Ds_Response>", "</Ds_Response>"
			GetParameterDiccionary data,"Ds_AuthorisationCode","<Ds_AuthorisationCode>","</Ds_AuthorisationCode>"
			GetParameterDiccionary data,"Ds_TransactionType","<Ds_TransactionType>", "</Ds_TransactionType>"
			GetParameterDiccionary data,"Ds_SecurePayment","<Ds_SecurePayment>", "</Ds_SecurePayment>"
			GetParameterDiccionary data,"Ds_Language","<Ds_Language>","</Ds_Language>"
			GetParameterDiccionary data,"Ds_MerchantData", "<Ds_MerchantData>", "</Ds_MerchantData>"
			GetParameterDiccionary data,"Ds_Card_Country","<Ds_Card_Country>","</Ds_Card_Country>"
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

	End Class

'    Dim key 
'    key="aggafccc121===bbabababeefee"
'    Dim test
'    Set test = New RedsysAPI
'    'dim c
'    'c=test.getOrderNotifSOAP("fhu 873qg<Response>Response=ZZZ</Response>h8f734<Ds_Order>Ds_Order=XXX</Ds_Order> dqhdniuq<Request>Request=YYY</Request>ediqwedqe")
'    'c=test.getRequestNotifSOAP("fhu 873qg<Response>Response=ZZZ</Response>h8f734<Ds_Order>Ds_Order=XXX</Ds_Order> dqhdniuq<Request>Request=YYY</Request>ediqwedqe")
'    'c=test.getResponseNotifSOAP("fhu 873qg<Response>Response=ZZZ</Response>h8f734<Ds_Order>Ds_Order=XXX</Ds_Order> dqhdniuq<Request>Request=YYY</Request>ediqwedqe")
'    test.SetParameter "aaaa","123"
'    test.SetParameter "bbbb","http://aaa.bbb.ccc:80/test.asp?vala=1&valb=2"
'    test.SetParameter "Ds_Merchant_Order","papapapas"
'    test.SetParameter "dddd","Son euros € no?"
'    test.SetParameter "DS_ORDER","papapapas"
'    dim d
'    d=test.createMerchantParameters()
'    dim sig1
'    sig1=test.createMerchantSignature(key)
'
'    Dim test2
'    Set test2 = New RedsysAPI
'    dim f
'    f=test2.decodeMerchantParameters(d)
'    dim sig2
'    sig2=test2.createMerchantSignatureNotif (key,d)
'stop
'    dim xmlReq
'    xmlReq = "<root><Request Ds_Version=""00""><Ds_Order>papapapas</Ds_Order><Fecha>2015/12/01</Fecha><Ds_Amount>100</Ds_Amount></Request><root>"
'    Dim test3
'    Set test3 = New RedsysAPI
'    dim sig3
'    sig3=test3.createMerchantSignatureNotifSOAPRequest(key,xmlReq)
'    test3.XMLToDiccionary xmlReq
'    xxx1=test3.GetParameter("fecha")
'    xxx1=test3.createMerchantParameters()
'    xx1=CryptoJS.enc.Base64.parse(xxx1).toString(CryptoJS.enc.Utf8)
'
'    dim xmlRes
'    xmlRes = "<root><Response Ds_Version=""00""><Ds_Order>papapapas</Ds_Order><Fecha>2015/12/01</Fecha><Ds_Amount>100</Ds_Amount></Response><root>"
'    Dim test4
'    Set test4 = New RedsysAPI
'    dim sig4
'    sig4=test4.createMerchantSignatureNotifSOAPResponse(key,xmlRes,"pedido002")
'    test4.XMLToDiccionary xmlRes
'
'    stop
'
</script>
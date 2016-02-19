<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>

<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/core.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/enc-base64.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/md5.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/evpkdf.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/cipher-core.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/tripledes.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/sha256.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/hmac.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-master-3.1.6/pad-zeropadding.js"></script>
<script type="text/javascript" language="javascript" runat="server">
	var CryptoJS; //Hay que declararla para que esté al alcance de VBScript 
</script>
<script language="vbscript" runat="server">
	Class Encrypt_Cfg
		public iv
		public mode
		public padding
		Function hasOwnProperty(name)
			hasOwnProperty = (name="iv" or name="mode" or name="padding")
		End Function
	End Class
</script>

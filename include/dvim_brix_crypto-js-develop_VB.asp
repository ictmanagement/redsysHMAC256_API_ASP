<script language="vbscript" runat="server">
'
'	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
'
</script>

<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/core.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/enc-base64.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/md5.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/evpkdf.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/cipher-core.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/tripledes.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/sha256.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/hmac.js"></script>
<script type="text/javascript" language="javascript" runat="server" src="./brix_crypto-js-develop-3.1.6/src/pad-zeropadding.js"></script>
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

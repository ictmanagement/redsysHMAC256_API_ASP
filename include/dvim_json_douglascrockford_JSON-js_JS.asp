<script type="text/javascript" language="javascript" runat="server">
/*
	Copyright (c) 2016, DE VEGA ICT MANAGEMENT, SLU, CIF B66216300
*/
</script>

<script src="./douglascrockford_JSON-js/json_parse.js" type="text/javascript" language="javascript" runat="server"></script>
<script type="text/javascript" language="javascript" runat="server">
	if (typeof JSON !== 'object') {
		JSON = {};
	}
	(function () {
		// Sobreescribimos el método parse original por el que no utiliza "eval"
		JSON.parse = json_parse;

		JSON.replacer = function (key, value) {
			// En caso que value sea un Dictionary, lo reemplazamos por un Array asociativo
			if (value && typeof value === 'object'  &&
				typeof value.Keys === 'unknown') {
				var ret = {};
				var keys = (new VBArray(value.Keys())).toArray();
				for (var k in keys)
					ret[keys[k]] = value.Item(keys[k]);
				return ret;
			}
			else {
				//return encodeURI(value);
				return value;
			}
		};
		JSON.reviver = function (key, value) {
			// En caso que value sea un Array asociativo, devolvemos un Dictionary 
			if (value && typeof value === 'object'
				&& Object.prototype.toString.apply(value) === '[object Object]'
				) {
				var ret = new ActiveXObject("Scripting.Dictionary");
				ret.CompareMode = 1;
				for (var k in value)
					ret.Item(k)=value[k];
				return ret;
			}
			else {
				return decodeURIComponent(value);
			}
		};
	}());
</script>
<script src="./douglascrockford_JSON-js/json2.js" type="text/javascript" language="javascript" runat="server"></script>
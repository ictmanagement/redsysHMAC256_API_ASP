# redsysHMAC256_API_ASP

Librerías de Ayuda (APIs) para integrar la pasarela de pago (TPV Virtual) Redsys en tiendas web virtuales que hayan sido desarrolladas bajo ASP Clásico, teniendo en cuenta el cambio del uso del algoritmo SHA1 por 3DES y HMAC-SHA256, que son más robustos (Ver https://canales.redsys.es/canales/ayuda/migracionSHA256.html)


Este API es un portado de las clases `RedSys` y `RedSysWs` implementadas en las API PHP que proporciona RedSys en su página http://www.redsys.es/wps/portal/redsys/publica/areadeserviciosweb/descargaDeDocumentacionYEjecutables. Por favor leer detenidamente las condiciones de uso de RedSys que podéis encontar en el enlace anterior, clicando en "Ver condiciones de uso" (no hay enlace directo).

## Dependencias

Para la implementación de las funciones HMAC-SHA256, 3DES y el manejo de los encoding Utf8, Base64 y Hex hemos utilizado la librería javascript https://github.com/brix/crypto-js de Evan Vosberg y Jeff Mott (@brix), para la que hemos creado librerías ASP que facilitan el acceso desde ASP Clásico en lenguaje VBScript ([dvim_brix_crypto-js-master_VB.asp](dvim_brix_crypto-js-master_VB.asp) o [dvim_brix_crypto-js-develop_VB.asp](dvim_brix_crypto-js-develop_VB.asp) según se utilize el branch [master](https://github.com/brix/crypto-js/tree/master) o [develop](https://github.com/brix/crypto-js/tree/develop) de https://github.com/brix/crypto-js. Esta librería está bajo licencia [MIT](http://opensource.org/licenses/MIT).

Para la implementación de JSON hemos utilizado la librería javascript https://github.com/douglascrockford/JSON-js de Douglas Crockford (@douglascrockford), en particular los ficheros [json2.js](https://github.com/douglascrockford/JSON-js/blob/master/json2.js) y [json_parse.js](https://github.com/douglascrockford/JSON-js/blob/master/json_parse.js), éste ultimo para usar una implementación de `parse` que no utilice `eval`. Hemos creado una librería ASP que facilita el acceso desde ASP Clásico en lenguaje VBScript, que además fuerza el uso de la implementación de `parse` de [json_parse.js](https://github.com/douglascrockford/JSON-js/blob/master/json_parse.js), y que implementa los métodos `replacer` y `reviver` (para gestionar el caso en que el `value` es un Dictionary). Esta librería es de Dominio Público, como se puede observar en el encabezado de cada fichero.

## Documentación
Para la utilización de estas APIs podéis seguir las indicaciones dadas por ResSys en su página https://canales.redsys.es/canales/ayuda/migracionSHA256.html, y en particular para el caso de PHP.
###Conexion Redirección
Para la migración de comercios existentes, descargar la [Guia de migración a HMAC SHA256 - conexion por redirección](https://canales.redsys.es/canales/ayuda/documentacion/Guia%20migracion%20a%20HMAC%20SHA256%20-%20conexion%20por%20redireccion.pdf).
Para el caso de nuevos comercios, descargar el [Manual integración para conexión por Redirección](https://canales.redsys.es/canales/ayuda/documentacion/Manual%20integracion%20para%20conexion%20por%20Redireccion.pdf)
###Conexion WebService
Para la migración de comercios existentes, descargar la [Guia de migración a HMAC SHA256 - conexión por Web Service](https://canales.redsys.es/canales/ayuda/documentacion/Guia%20migracion%20a%20HMAC%20SHA256%20-%20conexion%20por%20Web%20Service.pdf).
Para el caso de nuevos comercios, descargar el [Manual de integración para conexión por Web Service](https://canales.redsys.es/canales/ayuda/documentacion/Manual%20integracion%20para%20conexion%20por%20Web%20Service.pdf)

## Ficheros incluídos
* [include/dvim_apiRedsys_VB.asp](include/dvim_apiRedsys_VB.asp): Portado de `apiRedsys.php`, más algunas funciones de ayuda.
* [include/dvim_apiRedsysWs_VB.asp](include/dvim_apiRedsysWs_VB.asp): Portado de `apiRedsysWs.php`.
* [include/dvim_brix_crypto-js-develop_VB.asp](include/dvim_brix_crypto-js-develop_VB.asp): Para facilitar el acceso desde ASP Clásico en lenguaje VBScript a la librería https://github.com/brix/crypto-js en caso de utilizar el branch [develop](https://github.com/brix/crypto-js/tree/develop).
* [include/dvim_brix_crypto-js-master_VB.asp](include/dvim_brix_crypto-js-master_VB.asp): Para facilitar el acceso desde ASP Clásico en lenguaje VBScript a la librería https://github.com/brix/crypto-js en caso de utilizar el branch [master](https://github.com/brix/crypto-js/tree/master).
* [include/dvim_json_douglascrockford_JSON-js_JS.asp](include/dvim_json_douglascrockford_JSON-js_JS.asp): Para facilitar el acceso desde ASP Clásico en lenguaje VBScript a la librería https://github.com/douglascrockford/JSON-js.
* [include/brix_crypto-js-develop-3.1.6/](include/brix_crypto-js-develop-3.1.6): Carpeta donde colocar el branch [develop](https://github.com/brix/crypto-js/tree/develop), del que sólo son necesarios el fichero [LICENSE](https://github.com/brix/crypto-js/blob/develop/LICENSE) y la carpeta [src](https://github.com/brix/crypto-js/tree/develop/src).
* [include/brix_crypto-js-master-3.1.6/](include/brix_crypto-js-master-3.1.6): Carpeta donde colocar el branch [develop](https://github.com/brix/crypto-js/tree/develop), del que sólo son necesarios los ficheros [LICENSE](https://github.com/brix/crypto-js/blob/master/LICENSE) y los ficheros `.js` de la raíz.
* [include/douglascrockford_JSON-js/](include/douglascrockford_JSON-js): Carpeta donde colocar los ficheros [json2.js](https://github.com/douglascrockford/JSON-js/blob/master/json2.js) y [json_parse.js](https://github.com/douglascrockford/JSON-js/blob/master/json_parse.js).
* [ejemploGeneraPet.asp](ejemploGeneraPet.asp): Ejemplo de generación de petición con redirección a RedSys.
* [ejemploRecepcionaPet.asp](ejemploRecepcionaPet.asp): Ejemplo de recepción de notificación desde RedSys.
* [ejemploSOAP.asp](ejemploSOAP.asp): Ejemplo básico de recepción de notificación SOAP desde RedSys. Para la implementación de SOAP se sugiere el uso de librerías que implementen la validación WSDL y que generen y validen los encabezados SOAP.

## Licencia de uso
`New BSD` también llamada `BSD 3-clause`, ver [LICENSE](LICENSE).

## Soporte y contacto
Si necesitáis soporte en la migración de vuestra tienda en ASP Clásico, no dudad en contactarnos en el +34931767617 o enviando un email a migracionSHA256@ictmanagement.es.


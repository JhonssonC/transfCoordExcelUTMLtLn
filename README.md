# transfCoordExcelUTMLtLn

Codigo VBA (Excel) para transformar coordenadas UTM a LatLong y viceversa (Probado en la costa de Ecuador)

Prueba de Ejecucion:
![img](https://i.imgur.com/dpqTsTA.gif)

Para ejecutar: 
* Cree un archivo excel con una hoja especifica de la cual el codigo obtendra las referencias de cuales son las columnas que contienen coordenadas:

VAR

![Imgur1](https://i.imgur.com/eLyjWgp.png)

* Importe los modulos .bas .cls  (agadecimiento especial al post https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel) desde el editor de VBA de excel.

![Imgur2](https://i.imgur.com/doXrknC.png)

* En Excel construya en una hoja vacia la siguiente tabla poniendo especial atencion a las columnas especificadas en la hoja VAR en el paso anterior las columnas deben concordar con los encabezados, no textualmente pero si deben ser los datos que se especificaron el la hoja VAR.

![Imgur3](https://i.imgur.com/toouN3p.png)

Ejecutar la macro segun la necesidad y requerimiento.

Una vez la tabla tenga datos se puede ejecutar seleccionando uno a varios elementos de la columna CODIGO/ID (columna A), esto siempre que haya datos de referencia para realizar la transformacion, por ejemplo si necesito encontrar Latitud y Longitud debo tener X y Y y si necesito transformar a X y Y necesito tener Latitud y Longitud.

![img](https://i.imgur.com/dpqTsTA.gif)

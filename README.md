# transfCoordExcelUTMLtLn

VBA code (Excel) to transform UTM coordinates to LatLong and vice versa

Using web requests and responses made to the arcgisonline coordinate transformation service (json)

(Tested in Ecuador)


Execution Test:

![img](https://i.imgur.com/dpqTsTA.gif)

To execute:
* Create an excel file with a specific sheet from which the code will obtain the references of which columns contain coordinates:

VAR sheet

![Imgur1](https://i.imgur.com/AomEqDY.png)

* Import the .bas .cls modules from the excel VBA editor.
* Special thanks to the post https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel

![Imgur2](https://i.imgur.com/aSbpjgJ.png)

* In Excel, build the following table on an empty sheet, paying special attention to the columns specified in the VAR sheet in the previous step. The columns must match the headers, not textually, but they must be the data specified in the VAR sheet.

![Imgur3](https://i.imgur.com/toouN3p.png)

Execute the macro according to the need and requirement.

Once the table has data, it can be executed by selecting one or several elements from the CODE/ID column (column A), this as long as there is reference data to perform the transformation, for example if I need to find Latitude and Longitude I must have X and Y and if I need to transform to X and Y I need to have Latitude and Longitude.

![img](https://i.imgur.com/dpqTsTA.gif)


Bibliography:

https://utility.arcgisonline.com/arcgis/rest/services/Geometry/GeometryServer/project
https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel

# transfCoordExcelUTMLtLn

VBA code (Excel) to transform UTM coordinates to LatLong and vice versa (Tested on the coast of Ecuador)

Execution Test:

![img](https://i.imgur.com/dpqTsTA.gif)

to execute:
* create an excel file with a specific sheet from which the code will obtain the references of which are the columns that contain coordinates:

VAR

![Imgur1](https://i.imgur.com/eLyjWgp.png)

* import the .bas .cls modules (special thanks to the post https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel) from the excel VBA editor.

![Imgur2](https://i.imgur.com/doXrknC.png)

* In Excel, build the following table on an empty sheet, paying special attention to the columns specified in the VAR sheet in the previous step. The columns must match the headers, not textually, but they must be the data specified in the VAR sheet.

![Imgur3](https://i.imgur.com/toouN3p.png)

Execute the macro according to the need and requirement.

Once the table has data, it can be executed by selecting one or several elements from the CODE/ID column (column A), as long as there is reference data to perform the transformation, for example if I need to find Latitude and Longitude, I must have X and Y and if I need to transform to X and Y I need to have Latitude and Longitude.

![img](https://i.imgur.com/dpqTsTA.gif)

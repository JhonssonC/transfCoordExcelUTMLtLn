Attribute VB_Name = "TRANSFORM"
Sub FILTRO_TRANSFORMARCOORD_UTMTOLATLNG()

    If Selection.Cells.Rows.Count > 1 Then
        
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
        
        COMUMNA = Col_Letter(Selection.Cells.Column)
        
        Dim celda As Range
    
        For Each celda In Range(COMUMNA & RI & ":" & COMUMNA & RF).SpecialCells(xlCellTypeVisible)
            celda.Select
            If Range(Sheets("VAR").Range("B1") & ActiveCell.Row) <> "" And Range(Sheets("VAR").Range("B2") & ActiveCell.Row) <> "" Then
                UtmtoLatLong Range(Sheets("VAR").Range("B1") & ActiveCell.Row), Range(Sheets("VAR").Range("B2") & ActiveCell.Row), Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B4") & ActiveCell.Row)
            End If
        Next
    Else
        If Range(Sheets("VAR").Range("B19") & ActiveCell.Row) <> "" And Range(Sheets("VAR").Range("B20") & ActiveCell.Row) <> "" Then
            UtmtoLatLong Range(Sheets("VAR").Range("B1") & ActiveCell.Row), Range(Sheets("VAR").Range("B2") & ActiveCell.Row), Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B4") & ActiveCell.Row)
        End If
    End If

End Sub

Sub FILTRO_TRANSFORMARCOORD_LATLNGTOUTM()


    If Selection.Cells.Rows.Count > 1 Then
        
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
        
        COMUMNA = Col_Letter(Selection.Cells.Column)
        
        Dim celda As Range
    
        For Each celda In Range(COMUMNA & RI & ":" & COMUMNA & RF).SpecialCells(xlCellTypeVisible)
            celda.Select
            If Range(Sheets("VAR").Range("B3") & ActiveCell.Row) <> "" And Range(Sheets("VAR").Range("B4") & ActiveCell.Row) <> "" Then
                LatLongtoUtm Range(Sheets("VAR").Range("B4") & ActiveCell.Row), Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B1") & ActiveCell.Row), Range(Sheets("VAR").Range("B2") & ActiveCell.Row)
            End If
        Next
    Else
        If Range(Sheets("VAR").Range("B3") & ActiveCell.Row) <> "" And Range(Sheets("VAR").Range("B4") & ActiveCell.Row) <> "" Then
            LatLongtoUtm Range(Sheets("VAR").Range("B4") & ActiveCell.Row), Range(Sheets("VAR").Range("B3") & ActiveCell.Row), Range(Sheets("VAR").Range("B1") & ActiveCell.Row), Range(Sheets("VAR").Range("B2") & ActiveCell.Row)
        End If
    End If

End Sub


Private Function Col_Letter(lngCol As Variant)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Private Function LatLongtoUtm(X As Variant, Y As Variant, destinoX As Range, destinoY As Range)


    Dim strVar As String
    Dim clsJSON As clsJSParse
    
    strO = "https://utility.arcgisonline.com/arcgis/rest/services/Geometry/GeometryServer/project?f=json&outSR=32717&inSR=4326&geometries=%7B%22geometryType%22%3A%22esriGeometryPoint%22%2C%22geometries%22%3A%5B%7B%22x%22%3A" & Replace(X, ",", ".") & "%2C%22y%22%3A" & Replace(Y, ",", ".") & "%2C%22spatialReference%22%3A%7B%22wkid%22%3A4326%2C%22latestWkid%22%3A4326%7D%7D%5D%7D"

    result = requestWeb(strO)
        
    Set clsJSON = New clsJSParse
    
    strVar = result
    
    'special thanks
    'https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel
    
    clsJSON.LoadString strVar
            
    destinoX = clsJSON.Value(clsJSON.NumElements - 1)
    destinoY = clsJSON.Value(clsJSON.NumElements)
    

End Function

Private Function UtmtoLatLong(X As Variant, Y As Variant, destinoLong As Range, destinoLat As Range)
    
    Dim clsJSON As clsJSParse
    Dim strVar As String
    
    strO = "https://utility.arcgisonline.com/arcgis/rest/services/Geometry/GeometryServer/project?f=json&outSR=4326&inSR=32717&geometries=%7B%22geometryType%22%3A%22esriGeometryPoint%22%2C%22geometries%22%3A%5B%7B%22x%22%3A" & Replace(Format(X, "00000.000000000"), ",", ".") & "%2C%22y%22%3A" & Replace(Format(Y, "000000.000000000"), ",", ".") & "%2C%22spatialReference%22%3A%7B%22wkid%22%3A32717%2C%22latestWkid%22%3A32717%7D%7D%5D%7D"

    result = requestWeb(strO)
    
    Set clsJSON = New clsJSParse
    
    strVar = result

    'special thanks
    'https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel
    clsJSON.LoadString strVar
            
    destinoLat = clsJSON.Value(clsJSON.NumElements - 1)
    destinoLong = clsJSON.Value(clsJSON.NumElements)
    

End Function

Private Function requestWeb(STR As Variant)

Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    Dim arrResult() As Variant
    Dim myxml As String
   
    objHTTP.Open "GET", "" & STR, False

    objHTTP.SetRequestHeader "Content-Type", "text/json;charset=utf-8"
    objHTTP.Send
    requestWeb = objHTTP.responseText
    
End Function

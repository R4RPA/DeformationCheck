Sub AddBackgroundShape()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rectShp As Shape
    Dim shapeName As String
    Dim groupedShp As Shape

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            shapeName = shp.Name
            Set rectShp = ws.Shapes.AddShape(msoShapeRectangle, shp.Left, shp.Top, shp.Width, shp.Height)
            With rectShp
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Line.ForeColor.RGB = RGB(255, 255, 255)
                .ZOrder msoSendToBack
            End With
            rectShp.Name = "Rectangle_" & shapeName
            Set groupedShp = ws.Shapes.Range(Array(shapeName, rectShp.Name)).Group
            groupedShp.Name = shapeName
        Next shp
    Next ws
End Sub

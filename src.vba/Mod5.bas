Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("TextBox 14")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 4).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    Range("D6").Select
    ActiveSheet.Shapes.Range(Array("TextBox 13")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 4).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Transparency = 0
        .Solid
    End With
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoFalse
    Range("D6").Select
End Sub
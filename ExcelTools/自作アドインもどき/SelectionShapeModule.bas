Attribute VB_Name = "SelectionShapeModule"
Option Explicit

' -----------------------------------------------------
' Function ŠÖ”
' -----------------------------------------------------


' -----------------------------------------------------
' Sub ŠÖ”
' -----------------------------------------------------
Sub ‘I‘ğ‚µ‚½}‚É˜gü‚ğ•t‚¯‚é()
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
End Sub

Attribute VB_Name = "SelectionShapeModule"
Option Explicit

' -----------------------------------------------------
' Function 関数
' -----------------------------------------------------


' -----------------------------------------------------
' Sub 関数
' -----------------------------------------------------
Sub 選択した図に枠線を付ける()
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
End Sub

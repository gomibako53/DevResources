Attribute VB_Name = "SelectionShapeModule"
Option Explicit

' -----------------------------------------------------
' Function �֐�
' -----------------------------------------------------


' -----------------------------------------------------
' Sub �֐�
' -----------------------------------------------------
Sub �I�������}�ɘg����t����()
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
    End With
End Sub

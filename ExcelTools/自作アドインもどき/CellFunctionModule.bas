Attribute VB_Name = "CellFunctionModule"
Option Explicit

' -------------------------------
' このFunctionを直接ワークシートの数式から呼び出すことはできないので、
' このモジュールをワークシート側のモジュールにコピーして使用する
' -------------------------------

' セルの背景色を取得する
Public Function getCellColor(rng As Range) As Long
    getCellColor = rng.Interior.color
End Function

' 指定範囲のセルに指定の背景色のセルが存在するか判定
Public Function isExistCellColor(rng As Range, colorIndex As Long) As Boolean
    Dim r As Range
    For Each r In rng
    If r.Interior.color = colorIndex Then
        isExistCellColor = True
        Exit Function
    End If
    Next r
    isExistCellColor = False
End Function

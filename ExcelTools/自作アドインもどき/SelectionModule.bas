Attribute VB_Name = "SelectionModule"
Option Explicit

' -----------------------------------------------------
' Function 関数
' -----------------------------------------------------

Private Function func選択範囲の座標を取得(ByRef ref_rangeTop As Long, ByRef ref_rangeBottom As Long, _
                                        ByRef ref_rangeLeft As Long, ByRef ref_rangeRight As Long)
    ref_rangeTop = Selection(1).row
    ref_rangeBottom = Selection(Selection.count).row
    ref_rangeLeft = Selection(1).Column
    ref_rangeRight = Selection(Selection.count).Column
End Function

Private Function func表作成(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional titleLineBackColor As Long = -1, _
                            Optional titleLineLetterColor As Long = xlAutomatic, _
                            Optional titleLineNum As Long = 1)

    ' 全体の線描画
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeRight))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlHairline   ' 水平線は点線
    End With
    
    ' タイトル行の線描画
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
        .Interior.COLOR = titleLineBackColor
        .Font.FontStyle = "標準"
        .Font.ColorIndex = titleLineLetterColor
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' タイトル行が2行以上の場合は、中間の横ラインを出さない
    If titleLineNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
            .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        End With
    End If
End Function

Private Function func_枠作成(ByVal COLOR As Double, Optional frameType As Integer = 1)
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    Application.ScreenUpdating = False
    
    With Range(Cells(top, left), Cells(bottom, right))
    
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeLeft).Weight = xlHairline
        .Borders(xlEdgeRight).Weight = xlHairline
        
        If frameType = 1 Then
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        ElseIf frameType = 2 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlHairline   ' 水平線は点線
        End If
        
        .Interior.COLOR = COLOR
    End With
    
    Application.ScreenUpdating = True
End Function

' -----------------------------------------------------
' Sub 関数
' -----------------------------------------------------

Sub 表作成_濃紺()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 6299648              ' 濃紺
    ' タイトルの文字色
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' 白
    
    Call func表作成(top, bottom, left, right, BackColor, FontColor)
    
    ' 紺色フォーマットの場合の特殊設定
    With Range(Cells(top, left), Cells(top, right))
        .Font.FontStyle = "太字"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_濃紺_2行ver()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 6299648              ' 濃紺
    ' タイトルの文字色
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' 白
    
    Call func表作成(top, bottom, left, right, BackColor, FontColor, 2)
    
    ' 紺色フォーマットの場合の特殊設定
    With Range(Cells(top, left), Cells(top + 1, right))
        .Font.FontStyle = "太字"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_無色()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    Call func表作成(top, bottom, left, right)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_黄色()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 10092543              ' 黄色
    
    Call func表作成(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_オレンジ()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 10079487              ' オレンジ
    
    Call func表作成(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_緑()
    Application.ScreenUpdating = False
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 13434828              ' 緑
    
    Call func表作成(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 間の縦線を除去()
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub 枠_黄()
    func_枠作成 (10092543)
End Sub

Sub 枠2_黄()
    Call func_枠作成(10092543, 2)
End Sub

Sub 枠_グレー()
    func_枠作成 (12632256)
End Sub

Sub 行単位でセルの結合()
    Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' 行単位でセルの結合
    If (bottom - top) < 5000 Then   ' 選択しすぎている場合は処理が固まる可能性があるので何も処理させない。
        For i = top To bottom
            Range(Cells(i, left), Cells(i, right)).MergeCells = True
        Next i
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub 列単位でセルの結合()
    Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    ' 列単位でセルの結合
    For i = left To right
        Range(Cells(top, i), Cells(bottom, i)).MergeCells = True
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub ウィンドウ枠の固定をしなおす()
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

Sub 選択範囲の各セルを編集状態にしてEnter()
    ' ファイルパスが記載されたセルをいったん編集状態にしてEnterを押すと、ハイパーリンクがかかる。
    ' この関数は、ハイパーリンクをかけるために使用する

    Application.ScreenUpdating = False
    
    Dim i       As Long
    Dim c       As Range
    
    ' 選択されている範囲を取得
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func選択範囲の座標を取得(top, bottom, left, right)
    
    For Each c In Range(Cells(top, left), Cells(bottom, right))
        If c.Value <> "" Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        ElseIf VarType(c.Value) = vbError Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        End If
    Next c

    Application.ScreenUpdating = True
End Sub



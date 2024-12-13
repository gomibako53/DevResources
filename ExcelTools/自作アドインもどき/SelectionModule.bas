Attribute VB_Name = "SelectionModule"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' -----------------------------------------------------
' Function 関数
' -----------------------------------------------------

Private Function func選択範囲の座標を取得(ByRef ref_rangeTop As Long, ByRef ref_rangeBottom As Long, _
                                        ByRef ref_rangeLeft As Long, ByRef ref_rangeRight As Long)
    ref_rangeTop = Selection(1).Row
    ref_rangeBottom = Selection(Selection.count).Row
    ref_rangeLeft = Selection(1).Column
    ref_rangeRight = Selection(Selection.count).Column
End Function

Private Function func表作成(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional removeVerticalBorders As Boolean = False, _
                            Optional removeHorizontalBorders As Boolean = False, _
                            Optional titleLineBackColor As Long = -1, _
                            Optional titleLineLetterColor As Long = xlAutomatic, _
                            Optional titleLineNum As Long = 1)
    
    Dim i As Long, j As Long

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
        .Interior.color = titleLineBackColor
        .Font.FontStyle = "標準"
        .Font.colorIndex = titleLineLetterColor
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' タイトル行が2行以上の場合は、中間の横ラインを出さない
    If titleLineNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
            .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        End With
    End If
    
    ' 縦の罫線のラインが必要ない場合は消す
    If removeVerticalBorders Then
        For i = rangeLeft To rangeRight
            Dim emptyCount As Long
            emptyCount = Application.WorksheetFunction.CountIf(Range(Cells(rangeTop, i), Cells(rangeBottom, i)), "<>")
            If emptyCount <= 0 Then
                Range(Cells(rangeTop, i), Cells(rangeBottom, i)).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                Range(Cells(rangeTop, i), Cells(rangeBottom, i)).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
            End If
        Next i
    End If

    Application.ScreenUpdating = True
End Function

Private Function func表作成_横(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional titleColumnBackColor As Long = -1, _
                            Optional titleColumnLetterColor As Long = xlAutomatic, _
                            Optional titleColumnNum As Long = 1)

    ' 全体の線描画
    With Range(Cells(rangeToo, rangeLeft), Cells(rangeBottom, rangeRight))
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
        .Borders(xlInsideHorizontal).Weight = xlHairline    ' 水平線は点線
    End With

    ' タイトル列の線描画
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeLeft + titleColumnNum - 1))
        .Interior.color = titleColumnBackColor
        .Font.FontStyle = "標準"
        .Font.colorIndex = titleColumnLetterColor
        .Borders(xlEdgeRight).LineStyle = xlDouble
    End With

    ' タイトル列が2行以上の場合は、中間の縦ラインを出さない
    If titleColumnNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeLeft + titleColumnNum - 1))
            .Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        End With
    End If
    
    Application.ScreenUpdating = True
End Function

Private Function func_枠作成(ByVal color As Double, Optional frameType As Integer = 1)
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    With Range(Cells(Top, Left), Cells(bottom, Right))
    
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
            .Borders(xlInsideVertical).LineStyle = xlNone
        ElseIf frameType = 2 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlHairline   ' 水平線は点線
            .Borders(xlInsideVertical).LineStyle = xlNone
        End If
        
        .Interior.color = color
    End With
    
    Application.ScreenUpdating = True
End Function

' -----------------------------------------------------
' Sub 関数
' -----------------------------------------------------

Sub 表作成_濃紺()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 6299648              ' 濃紺
    ' タイトルの文字色
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' 白
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor, FontColor)
    
    ' 紺色フォーマットの場合の特殊設定
    With Range(Cells(Top, Left), Cells(Top, Right))
        .Font.FontStyle = "太字"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_濃紺_2行ver()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 6299648              ' 濃紺
    ' タイトルの文字色
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' 白
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor, FontColor, 2)
    
    ' 紺色フォーマットの場合の特殊設定
    With Range(Cells(Top, Left), Cells(Top + 1, Right))
        .Font.FontStyle = "太字"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_無色()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    Call func表作成(Top, bottom, Left, Right)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_無色_2行()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    Call func表作成(Top, bottom, Left, Right, True, True, -1, xlAutomatic, 2)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_無色_横()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    Call func表作成_横(Top, bottom, Left, Right)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_黄色()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 10092543              ' 黄色
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_オレンジ()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 10079487              ' オレンジ
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_緑()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 13434828              ' 緑
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_グレー()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 15395562              ' 薄めのグレー
    
    Call func表作成(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 表作成_グレー_横()
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' タイトルの背景色
    Dim BackColor As Long: BackColor = 15395562              ' 薄めのグレー
    
    Call func表作成_横(Top, bottom, Left, Right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub 枠_黄()
    func_枠作成 (10092543)
End Sub

Sub 枠2_黄()
    Call func_枠作成(10092543, 2)
End Sub

Sub 枠_グレー()
    func_枠作成 (15395562)  ' 薄めのグレー
End Sub

Sub 列単位でセルの結合()
    'Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    ' 列単位でセルの結合
    For i = Left To Right
        Range(Cells(Top, i), Cells(bottom, i)).MergeCells = True
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

    'Application.ScreenUpdating = False
    
    Dim i       As Long
    Dim c       As Range
    
    ' 選択されている範囲を取得
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)
    
    For Each c In Range(Cells(Top, Left), Cells(bottom, Right))
        If c.HasFormula Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        ElseIf c.Value <> "" Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        End If
        Sleep (500)
    Next c

    Application.ScreenUpdating = True
End Sub

Sub 選択したセルのコメント位置を修正()
    Dim targetRange As Range: Set targetRange = Selection
    Dim myRange As Range

    For Each myRange In targetRange
        If Not (myRange.Comment Is Nothing) Then
            With myRange.Comment.shape
                .Top = myRange.Top
                .Left = myRange.Offset(, 1).Left
                .TextFrame.AutoSize = True
            End With
        End If
    Next
End Sub

Sub 選択範囲でブランクのセルは上のセル値を入れる()
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Dim i As Long, j As Long

    'Application.ScreenUpdating = False

    ' 選択されている範囲を取得
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)

    For i = Left To Right
        For j = Top To bottom
            ' マージされているセルは無視
            If Not Cells(j, i).MergeCells Then
                ' 非表示のセルは無視
                If Not Cells(j, i).Rows.Hidden Then
                    If Not Cells(j, i).Columns.Hidden Then
                        ' 値が入っていなかったらすぐ上のセルの値と同じものを入れる
                        If j >= 2 And Len(Cells(j, i).Text) = 0 And Len(Cells(j - 1, i).Text) > 0 Then
                            Cells(j, i).Value = Cells(j - 1, i).Value
                        End If
                    End If
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub

Sub 選択範囲で上のセル値と異なる場合はセル色を黄色にする()
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Dim i As Long, j As Long

    ' 選択されている範囲を取得
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)

    For i = Left To Right
        For j = Top To bottom
            ' マージされているセルは無視
            If Not Cells(j, i).MergeCells Then
                ' 値が異なったら色を付ける
                If j >= 2 And Len(Cells(j, i).Text) > 0 And Cells(j - 1, i).Text <> Cells(j, i).Text Then
                    Cells(j, i).Interior.color = 10092543  ' 黄色
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub

Sub 選択範囲で上のセルと同じならフォント色をグレーにする()
    Dim selectedRange As Range
    Dim cell As Range
    Dim previousCell As Range

    ' 選択範囲を取得
    Set selectedRange = Selection

    ' 2行目以降の各セルに対して処理を実行
    For Each cell In selectedRange
        If cell.Row > 1 Then  ' 1行目は無視する
            Set previousCell = cell.Offset(-1)   ' 上のセルを取得

            ' セルがブランクでない場合に処理を実行
            If Not IsEmpty(cell.Value) Then
                ' セルの値が上のセルと同じか比較
                If cell.Value = previousCell.Value Then
                    ' セルが数式で表されている場合は計算結果を比較
                    If cell.HasFormula Then
                        If cell.Value = previousCell.Value Then
                            cell.Font.color = RGB(192, 192, 192)   ' グレーにする
                        End If
                    Else
                        cell.Font.color = RGB(192, 192, 192)    ' グレーにする
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Sub 選択範囲で上のセルと同じならフォント色を薄くする()
    Dim selectedRange As Range
    Dim cell As Range
    Dim previousCell As Range
    Dim currentColor As Long
    Dim newColor As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    ' 選択範囲を取得
    Set selectedRange = Selection

    ' 2行目以降の各セルに対して処理を実行
    For Each cell In selectedRange
        If cell.Row > 1 Then  ' 1行目は無視する
            Set previousCell = cell.Offset(-1)   ' 上のセルを取得

            ' セルがブランクでない場合に処理を実行
            If Not IsEmpty(cell.Value) Then
                ' セルの値が上のセルと同じか比較
                If cell.Value = previousCell.Value Then
                    ' セルが数式で表されている場合は計算結果を比較
                    If cell.HasFormula Then
                        If cell.Value = previousCell.Value Then
                            currentColor = cell.Font.color ' 現在の色を取得
                            redValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - (currentColor And 255)) * 3 / 4 + (currentColor And 255), 0), 255) ' 赤の値を計算
                            greenValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256) And 255)) * 3 / 4 + ((currentColor \ 256) And 255), 0), 255)  ' 緑の値を計算
                            blueValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256 \ 256) And 255)) * 3 / 4 + ((currentColor \ 256 \ 256) And 255), 0), 255) ' 青の値を計算
                            newColor = RGB(redValue, greenValue, blueValue) ' 薄い色を計算
                            cell.Font.color = newColor   ' 色を設定
                        End If
                    Else
                        currentColor = cell.Font.color ' 現在の色を取得
                        redValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - (currentColor And 255)) * 3 / 4 + (currentColor And 255), 0), 255) ' 赤の値を計算
                        greenValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256) And 255)) * 3 / 4 + ((currentColor \ 256) And 255), 0), 255)  ' 緑の値を計算
                        blueValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256 \ 256) And 255)) * 3 / 4 + ((currentColor \ 256 \ 256) And 255), 0), 255) ' 青の値を計算
                        newColor = RGB(redValue, greenValue, blueValue) ' 薄い色を計算
                        cell.Font.color = newColor   ' 色を設定
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Sub 選択範囲で星マークのセルが無い列は折りたたみ設定()
    Dim myRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    For Each myRange In Selection
        If myRange.Value = "★" Or myRange.Value = "☆" Then
            If Not map.Exists(myRange.Column) Then
                Call map.Add(myRange.Column, True)
            End If
        End If
    Next

    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub 選択範囲で星マークのセルの下は赤文字()
    Dim myRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    For Each myRange In Selection
        If myRange.Value = "★" Or myRange.Value = "☆" Then
            Cells(myRange.Row + 1, myRange.Column).Font.color = 255
        End If
    Next
End Sub

Sub 選択範囲で値が入っていない列は折りたたみ設定()
    Dim myRange As Range
    Dim targetRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)

    If Top = bottom Then
        ' 1行しか選択されていない場合は素直にその範囲だけ処理する
        Set targetRange = Selection
    Else
        ' 複数行選択されている場合は1行目は無視する。このマクロは表形式で使われることを想定していて、ヘッダ行は無視したいので。
        Set targetRange = Range(Cells(Top + 1, Left), Cells(bottom, Right))
    End If

    ' 値が入っている列を記憶
    For Each myRange In targetRange
        If Not IsEmpty(myRange.Value) Then
            If Not map.Exists(myRange.Column) Then
                Call map.Add(myRange.Column, True)
            End If
        End If
    Next

    ' 値が入っていない列は折りたたみ
    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub 選択範囲で値が入っていない列は折りたたみ設定_塗りつぶしのあるセルは除外()
    Dim myRange As Range
    Dim targetRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func選択範囲の座標を取得(Top, bottom, Left, Right)

    If Top = bottom Then
        ' 1行しか選択されていない場合は素直にその範囲だけ処理する
        Set targetRange = Selection
    Else
        ' 複数行選択されている場合は1行目は無視する。このマクロは表形式で使われることを想定していて、ヘッダ行は無視したいので。
        Set targetRange = Range(Cells(Top + 1, Left), Cells(bottom, Right))
    End If

    ' 値が入っている列を記憶
    For Each myRange In targetRange
        If Not IsEmpty(myRange.Value2) Then
            ' 塗りつぶしされていなかった場合のみ
            If myRange.Interior.colorIndex = xlNone Then
                If Not map.Exists(myRange.Column) Then
                    Call map.Add(myRange.Column, True)
                End If
            End If
        End If
    Next

    ' 値が入っていない列は折りたたみ
    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub 選択中のセル位置をクリップボードにコピー()
    Dim address As String
    Dim sheetName As String
    Dim msg As String
    address = Selection.address
    address = Replace(address, "$", "")
    sheetName = ActiveSheet.Name

    msg = ActiveWorkbook.Name & vbLf & "'" & sheetName & "'!" & address

    PutClipBoard (msg)
    Application.StatusBar = msg  ' デバッグ用に一時的に書いた。
End Sub

' 選択範囲のセルの書式を更新するマクロ
' 選択範囲の中で一番左上のセルの以下の書式を他の選択セルに反映する
'   ・セルの色(背景色)
'   ・文字色
'   ・罫線
'   ・書式体
Sub 左上セルの書式を他セルにコピー()
    Dim sourceCell As Range
    Dim targetRange As Range

    ' 選択範囲の中で一番左上のセルを取得
    Set sourceCell = Selection.Cells(1)

    ' 選択範囲を取得
    Set targetRange = Selection

    ' 書式をコピーする
    sourceCell.Copy

    ' 書式を他の選択セルに反映する
    targetRange.PasteSpecial xlPasteFormats

    ' ステータスバーにメッセージを表示し、5秒後に消す
    Application.StatusBar = "書式をコピーしました"
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Sub 選択範囲の横幅を自動調整()
    ' 列の幅の自動調節 (選択範囲で調節)
    Selection.Columns.AutoFit

    ' 選択範囲の列をループして、横幅が大きすぎるものは固定長に直す
    Const MAX_WIDTH As Double = 50 ' 最大の列幅を定義
    Dim col As Range
    For Each col In Selection.Columns
        If col.ColumnWidth > MAX_WIDTH Then
            col.ColumnWidth = MAX_WIDTH
        End If
    Next col

    Application.ScreenUpdating = True
End Sub

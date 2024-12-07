Attribute VB_Name = "BatchModule"
Option Explicit

Sub シート上のオブジェクトを全削除()
    Dim rc As Long
    rc = MsgBox("Are you sure to delete all shapes?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If
    
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub

Sub カレントシートのコメントを全削除()
    Dim rc As Long
    rc = MsgBox("Are you sure to delete all comments?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If
    
    On Error Resume Next
    Cells.SpecialCells(xlCellTypeComments).ClearComments
End Sub

Sub アクティブブックのオートフィルタを全て解除する()
    Dim sh As Worksheet
    
    For Each sh In Worksheets
        If sh.AutoFilterMode Then
            If sh.AutoFilter.FilterMode Then
                sh.ShowAllData
            End If
        End If
    Next sh
End Sub

Sub B2セルの内容でシート名を変更()
    Dim str As String: str = ActiveSheet.Cells(2, 2).Value
    Dim length As Long: length = Len(str)
    
    If length > O And length <= 31 Then
        ActiveSheet.Name = str
    End If
End Sub


Sub カレントBookの名前の定義を削除()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' エラーを無視
        n.Delete
    Next
End Sub

Sub カレントBookの名前の定義を削除_印刷設定以外()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If InStr(n.Name, "Print_") = 0 Then
            On Error Resume Next ' エラーを無視
            n.Delete
        End If
    Next
End Sub

Sub カレントBookの名前の定義を削除_エラーのみ()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' エラーを無視
               If InStr(n.Value, "=#") = 1 Then
            n.Delete
        End If
    Next
End Sub

Sub カレントBookの名前の定義で別ブックの参照をしているものを削除()
    Dim n As Name
    Dim count As Long: count = 0

    Application.StatusBar = False

    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' エラーを無視

        If left(n.Value, 4) = "='\\" Or left(n.Value, 5) = "='C:\" Then
            count = count + 1
            n.Delete
        End If
    Next

    ' ステータスパの更新
    If count <> 0 Then
        Application.StatusBar "別ブック参照の名前：" & count & "件を削除しました。"
    Else
        Application.StatusBar = False
    End If
End Sub

Sub DIFFガ式のシートを整形()
    Dim activeSheetBak As Worksheet: Set activeSheetBak = ActiveSheet
    Dim sh As Worksheet

    Application.ScreenUpdating = False

    For Each sh In Worksheets
        ' A1セルがDiffツールで使っている青色だったらDiffシートと判断
        If sh.Range("A1").Interior.COLOR = 16711680 Then
            Call func_DIFF形式のシートを整形_1シート(sh)
        End If
    Next sh

    activeSheetBak.Activate
    Application.ScreenUpdating = True
End Sub

Sub カレントBookの名前の定義で別ブック参照がある箇所を検出()
    Dim n As Name
    Dim jumpedFlg As Boolean: jumpedFlg = False
    Dim count As Long: count = 0
    Dim firstHit As String

    Application.StatusBar = False

    For Each n In ActiveWorkbook.Names
        On Error Resume Next 'エラーを無視

        If left(n.Value, 4) = "='\\" Or left(n.Value, 5) = "='C:\" Then
            count = count + 1
            Debug.Print "--------------------------"
            Debug.Print n.Name
            Debug.Print n.Value
            Debug.Print n.Parent.CodeName
            Debug.Print n.Parent.Authore

            ' 最初にヒットした名前にジャンプする
            If jumpedFlg = False Then
                Application.GoTo Reference:=n.Name
                jumpedFlg = True
                firstHit = "name:[" & n.Name & "] CodeName:[" & n.Parent.CodeName & "]"
            End If
        End If
    Next

    ' ステータスバーの更新
    If jumpedFlg Then
        Application.StatusBar = "別ブック参照の名前が" & count & "件見つかりました。(first Hit ->" & firstHit & ")"
    Else
        Application.StatusBar = False
    End If
End Sub

Private Function func_DIFF形式のシートを整形_1シート(ByVal sh As Worksheet)
    sh.Select
    ' ウィンドウ枠の固定
    sh.Rows("2:2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' 表示サイズを75%に
    ActiveWindow.Zoom = 75

    ' 行番号が書いてあるA列とC列のサイズを調整
    sh.Range("A:A,C:C").ColumnWidth = 5

    ' ファイルの中身が書いてあるB列とD列のサイズを調整
    sh.Range("B:B,D:D").ColumnWidth = 95

    ' DIFF列の設定、調整
    sh.Range("E1").FormulaR1C1 = "DIFF"
    sh.Columns("E:E").ColumnWidth = 4

    ' 既にオートフィルタが設定されている場合は解除
    If sh.AutoFilterMode = True Then
        If sh.AutoFilter.FilterMode = True Then
            sh.ShowAllData
        End If
        sh.Rows("1:1").AutoFilter
    End If

    ' オート2フィルタの設定
    sh.Rows("1:1").AutoFilter

    ' 差分のある行(黄色の行)をフィルタ
    sh.Range(Range("A1").Cells(Rows.count, 3).End(xlUp)).AutoFilter Field:=4, Criteria1:=RGB(239, 203, 5), Operator:=xlFilterCellColor
    
    ' E1セルを選択
    sh.Range("E1").Select
End Function



Sub シート状のオブジェクトを全削除()
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub



Sub ワークシートを名前でソート()
    Application.ScreenUpdating = False
    
    Dim i As Long, j As Long, cnt As Long
    Dim buf() As String, swap As String
    Dim selectedSheet As Worksheet
    
    ' 元々開いていたシートを記憶
    Set selectedSheet = Application.ActiveWorkbook.ActiveSheet
    
    cnt = Application.ActiveWorkbook.Worksheets.count
    
    If cnt > 1 Then
        ReDim buf(cnt)
        
        ' ワークシート名を配列に入れる
        For i = 1 To cnt
            buf(i) = Application.ActiveWorkbook.Worksheets(i).Name
        Next i
        
        ' 配列の要素をソートする
        For i = 1 To cnt
            For j = cnt To i Step -1
                If buf(i) > buf(j) Then
                    swap = buf(i)
                    buf(i) = buf(j)
                    buf(j) = swap
                End If
            Next j
        Next i
    End If
    
    ' ワークシートの位置を並び替える
    Application.ActiveWorkbook.Worksheets(buf(1)).Move Before:=Application.ActiveWorkbook.Worksheets(1)
    
    For i = 2 To cnt
        Application.ActiveWorkbook.Worksheets(buf(i)).Move After:=Application.ActiveWorkbook.Worksheets(i - 1)
    Next i
    
    selectedSheet.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub シート名一覧を作成する()
    Application.ScreenUpdating = False
    
    Const addSheetName As String = "シート名一覧(自動生成)"
    
    Dim i As Long
    Dim ws As Worksheet
    Dim flag As Boolean
    
    For Each ws In Application.ActiveWorkbook.Worksheets
        ' addSheetNameの名称のシートが見つかったら削除する
        If ws.Name = addSheetName Then
            Application.DisplayAlerts = False   ' 削除時の警告メッセージは非表示
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
        
    Application.ActiveWorkbook.Sheets.Add Before:=Application.ActiveWorkbook.Worksheets(1)
    Application.ActiveWorkbook.ActiveSheet.Name = addSheetName
    
    For i = 2 To Application.ActiveWorkbook.Sheets.count
        Application.ActiveWorkbook.ActiveSheet.Cells(i - 1, "A").Value = Application.ActiveWorkbook.Sheets(i).Name
    Next i

    Application.ScreenUpdating = True
End Sub

Attribute VB_Name = "BatchModule"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub シート上のオブジェクトを全削除()
    Dim rc As Integer
    rc = MsgBox("Are you sure to delete all shapes on current sheet?", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
        ActiveSheet.Shapes.SelectAll
        Selection.Delete
    End If
End Sub

Sub シート上のオブジェクトの書式変更_セルにあわせて移動やリサイズをする()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMoveAndSize ' セル削除や移動に合わせて移動し、リサイズも行う
End Sub

Sub シート上のオブジェクトの書式変更_セルにあわせて移動やリサイズしない()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlFreeFloating ' セル削除や移動に合わせたリサイズ、移動を行わない
End Sub

Sub シート上のオブジェクトの書式変更_セルにあわせて移動するがリサイズしない()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMove ' セル削除や移動に合わせて移動する
End Sub

Sub カレントシートのコメントを全削除()
    Dim rc As Integer
    rc = MsgBox("Are you sure to delete all comments on current sheet?", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
        On Error Resume Next
        Cells.SpecialCells(xlCellTypeComments).ClearComments
    End If
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
    Dim newSheetName As String
    Dim temp As String
    Dim i As Integer
    
    If length > 0 And length <= 31 Then
        newSheetName = str
    ElseIf length > 31 Then
        newSheetName = left(str, 31)
    End If

    If ActiveSheet.Name <> newSheetName And SheetExists(newSheetName) Then
        If Len(newSheetName) > 28 Then
            newSheetName = left(newSheetName, 28)
        End If
        
        For i = 2 To 10
            temp = newSheetName & "(" & i & ")"
            If ActiveSheet.Name = temp Or Not SheetExists(temp) Then
                newSheetName = temp
                Exit For
            End If
        Next i
    End If
    
    If Len(newSheetName) > 0 Then
        ActiveSheet.Name = newSheetName
    End If
End Sub

Function SheetExists(shtName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(shtName) Is Nothing
    On Error GoTo 0
End Function


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
        Application.StatusBar = "別ブック参照の名前：" & count & "件を削除しました。"
    Else
        Application.StatusBar = False
    End If
End Sub

Sub DIFF方式のシートを整形()
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

Sub DIFF方式のシートを整形_アクティブシートのみ()
    Application.ScreenUpdating = False
    Call func_DIFF形式のシートを整形_1シート(ActiveSheet)
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
                Application.Goto Reference:=n.Name
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
    Dim YELLOW As Long: YELLOW = RGB(239, 203, 5)
    Dim GRAY As Long: GRAY = RGB(192, 192, 192)
    Dim LIGHT_PINK As Long: LIGHT_PINK = RGB(240, 192, 192)
    Dim PINK As Long: PINK = RGB(239, 119, 116)
    
    Dim i As Long
    Dim sheetName As String
    Dim bottom As Long
    Dim tmp As Long

    sh.Select
    ' ウィンドウ枠の固定
    sh.Rows("2:2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' 表示サイズを変更
    ActiveWindow.Zoom = 85

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
    bottom = Cells(Rows.count, 1).End(xlUp).Row   ' A列の最終行
    tmp = Cells(Rows.count, 3).End(xlUp).Row ' C列の最終行
    If bottom < tmp Then
        bottom = tmp
    End If
    For i = 2 To bottom
        If Cells(i, 2).Interior.COLOR = GRAY Or Cells(i, 4).Interior.COLOR = GRAY Or _
           Cells(i, 2).Interior.COLOR = YELLOW Or Cells(i, 4).Interior.COLOR = YELLOW Or _
           Cells(i, 2).Interior.COLOR = LIGHT_PINK Or Cells(i, 4).Interior.COLOR = LIGHT_PINK Or _
           Cells(i, 2).Interior.COLOR = PINK Or Cells(i, 4).Interior.COLOR = PINK Then
            ' 直前の行が差分ありと判定されていたら、次のElseIfの判定をせずにこの行も差分ありと判定。
            ' 連続して一塊になっているなら、この行だけ差分無しと判定しても使い勝手が悪いので。
            If Cells(i - 1, 5).Value = "★" Then
                Cells(i, 5).Value = "★"
            ' 空行は除外。あと、importで始まる行も除外。
            ElseIf Not (Len(Cells(i, 2).Value) = 0 And Len(Cells(i, 4).Value) = 0) And _
              Not (CommonModule.regularExpressionTest(Cells(i, 2).Value, "^import .+;", False)) And _
              Not (CommonModule.regularExpressionTest(Cells(i, 4).Value, "^import .+;", False)) Then
                Cells(i, 5).Value = "★"
           End If
        End If
    Next i
    
    ' E1セルを選択
    sh.Range("E1").Select
End Function

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

Sub シート名一覧をクリップボードにコピー()
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetNames As String: sheetNames = ""

    For Each ws In Application.ActiveWorkbook.Worksheets
        sheetNames = sheetNames & ws.Name & vbCrLf
    Next ws
    
    Call PutClipBoard(sheetNames)
End Sub

Sub 全シート倍率100パーセントにして先頭セル選択()
    Dim sht   As Worksheet              ' 処理中のワークシート
    Dim shtVisible                      ' 表示可能なワークシート
    Dim iRow, iCol                      ' 縦、横座標
    Dim oFilterStatus As AutoFilter     ' オートフィルタ状態
    Dim oRangeFilter As Range           ' オートフィルタ設定
    Dim zoomRc As Integer
    Dim zoomMsgBoxConducted As Boolean: zoomMsgBoxConducted = False

    Application.ScreenUpdating = True

    For Each sht In Sheets
        If (IsEmpty(shtVisible) = True) And (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            Set shtVisible = sht
        End If

        ' シートが表示されている場合
        If (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            sht.Select

            ' 85%以下のときは85%にする
            If ActiveWindow.Zoom <= 85 Then
                ActiveWindow.Zoom = 85
            Else
                ActiveWindow.Zoom = 100
            End If

            ' ウインドウ枠の固定がされている場合
            If ActiveWindow.FreezePanes = True Then
                iRow = ActiveWindow.SplitRow + 1
                iCol = ActiveWindow.SplitColumn + 1
                Cells(iRow + 1, iCol + 1).Activate
            End If

            Set oFilterStatus = sht.AutoFilter
            ' オートフィルタが設定されている場合
            If Not oFilterStatus Is Nothing Then
                ' フィルタが掛かっている場合
                If oFilterStatus.FilterMode = True Then
                    ' フィルタが掛かっている行の先頭を選択
                    Set oRangeFilter = Range("A1").CurrentRegion
                    Set oRangeFilter = Application.Intersect(oRangeFilter, oRangeFilter.Offset(1, 0))
                    Set oRangeFilter = oRangeFilter.SpecialCells(xlCellTypeVisible)
                    Range("A" & CStr(oRangeFilter.Row)).Select
                End If
            End If
            
            sht.Range("A1").Select
        End If
    Next
    
    shtVisible.Select

End Sub

Sub 全シート改ページプレビュー_枠線無()
    Dim sht As Worksheet ' 処理中のワークシート
    Dim shtVisible      ' 表示可能なワークシート
    Dim iRow, iCol ' 縦、横座標
    Dim oFilterStatus As AutoFilter  ' オートフィルタ状態
    Dim oRangeFilter As Range ' オートフィルタ設定
    Dim zoomRc As Integer
    Dim zoomMsgBoxConducted As Boolean: zoomMsgBoxConducted = False
    Dim rc As Integer
    
    rc = MsgBox("Are you sure to change all sheets format ?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = True
    
    For Each sht In Sheets
        If (IsEmpty(shtVisible) = True) And (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            Set shtVisible = sht
        End If

        ' シートが表示されている場合
        If (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            sht.Select
            ActiveWindow.View = xlPageBreakPreview ' 改ページプレビュー
            
            ' 85%以下のときは85%にする
            If ActiveWindow.Zoom <= 85 Then
                ActiveWindow.Zoom = 85
            Else
                ActiveWindow.Zoom = 100
            End If
            ActiveWindow.DisplayGridlines = False ' 枠線無し

            ' ウインドウ枠の固定がされている場合
            If ActiveWindow.FreezePanes = True Then
                iRow = ActiveWindow.SplitRow + 1
                iCol = ActiveWindow.SplitColumn + 1
                Cells(iRow + 1, iCol + 1).Activate
            End If

            Set oFilterStatus = sht.AutoFilter
            ' オートフィルタが設定されている場合
            If Not oFilterStatus Is Nothing Then
                ' フィルタが掛かっている場合
                If oFilterStatus.FilterMode = True Then
                    ' フィルタが掛かっている行の先頭を選択
                    Set oRangeFilter = Range("A1").CurrentRegion
                    Set oRangeFilter = Application.Intersect(oRangeFilter, oRangeFilter.Offset(1, 0))
                    Set oRangeFilter = oRangeFilter.SpecialCells(xlCellTypeVisible)
                    Range("A" & CStr(oRangeFilter.Row)).Select
                End If
            End If

            sht.Range("A1").Select
        Else
            sht.Visible = xlSheetVisible ' シートを表示
            sht.Select
            
            ActiveWindow.DisplayGridlines = False ' 枠線無し
            
            sht.Visible = xlSheetHidden ' シートを非表示
        End If
    Next

    shtVisible.Select

End Sub

Sub 全シート指定倍率に変更()
    Dim shOriginalSelected As Worksheet
    Dim sh As Worksheet
    Dim strScale As String
    Dim nScale As Long
    
    strScale = InputBox("倍率を指定してください(""**%""の数字部分を入力)")
    If strScale = "" Then Exit Sub
    nScale = CLng(strScale)

    Application.ScreenUpdating = False
    
    Set shOriginalSelected = ActiveSheet

    For Each sh In Sheets
        If (sh.Visible <> xlSheetHidden) And (sh.Visible <> xlSheetVeryHidden) Then
            sh.Select
            ActiveWindow.Zoom = nScale
        End If
    Next
    
    shOriginalSelected.Activate

    Application.ScreenUpdating = True
End Sub

Sub クリップボードの解放()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
    'Application.CutCopyMode = False
End Sub

Sub カレントシートの全セルをMeiryoUIに()
    Cells.Font.Name = "Meiryo UI"
End Sub

Sub シート表示設定_色が付いてないシートは非表示に()
    Dim ws As Worksheet
    Dim containColorTab As Boolean
    Application.ScreenUpdating = False
    
    containColorTab = False
    For Each ws In Worksheets
        If ws.Tab.ColorIndex <> xlNone Then
            containColorTab = True
            Exit For
        End If
    Next ws
    
    If containColorTab Then
        For Each ws In Worksheets
            If ws.Tab.ColorIndex = xlNone Then
                If ws.Visible = xlSheetVisible Then
                    ws.Visible = False
                End If
            End If
        Next ws
        Application.StatusBar = False
    Else
        Application.StatusBar = "色付きシートが無いので処理しませんでした"
    End If

    Application.ScreenUpdating = True
End Sub

Sub シート表示設定_全シート表示()
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each ws In Worksheets
        ws.Visible = True
    Next ws
    Application.ScreenUpdating = True

    Application.StatusBar = False
End Sub

Sub アクティブブックの印刷方向を全て横に()
    Dim sh As Worksheet
    Dim rc As Integer
    Dim preCheckResult As String

    ' まずは横向きと縦向きのどちらになっているかチェック
    For Each sh In Worksheets
        ' 横向き
        If sh.PageSetup.Orientation = xlLandscape Then
            preCheckResult = preCheckResult & "Landscape"
        ' 縦向き
        Else
            preCheckResult = preCheckResult & "Portrait"
        End If
        
        preCheckResult = preCheckResult & " : " & sh.Name & vbCrLf
    Next sh
    
    rc = MsgBox(preCheckResult & vbCrLf & "Are you sure to set LANDSCAPE mode on all sheets?", vbYesNo + vbQuestion)

    If rc = vbYes Then
        For Each sh In Worksheets
            If sh.PageSetup.Orientation = xlPortrait Then
                sh.PageSetup.Orientation = xlLandscape
            End If
        Next sh
    End If
End Sub

Sub アクティブブックをCSV形式で保存()
    Application.ScreenUpdating = False

    Dim cnt As Long: cnt = 1
    Dim bookName As String
    Dim directoryPath As String
    Dim activeBookFullPath As String

    bookName = left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    activeBookFullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name

    directoryPath = ActiveWorkbook.Path & "\" & bookName
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
    End If

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveWorkbook.SaveAs Filename:=directoryPath & "\" & cnt & "_" & ws.Name & ".csv", FileFormat:=xlCSV
        cnt = cnt + 1
    Next ws
    
    ' アクティブブックがCSV形式の名前になっているので、一度ファイルを閉じて再度開く。
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open activeBookFullPath

    Application.ScreenUpdating = True
    MsgBox "アクティブブックをCSV形式で保存しました"
End Sub

Sub アクティブシートをCSV形式で保存()
    Application.ScreenUpdating = False
    
    Dim bookName As String
    Dim directoryPath As String
    Dim activeBookFullPath As String
    
    bookName = left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    activeBookFullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name

    directoryPath = ActiveWorkbook.Path & "\" & bookName
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
    End If

    ActiveWorkbook.SaveAs Filename:=directoryPath & "\" & ActiveSheet.Name & ".csv", FileFormat:=xlCSV
    
    ' アクティブブックがCSV形式の名前になっているので、一度ファイルを閉じて再度開く。
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open activeBookFullPath
    
    Application.ScreenUpdating = True
    MsgBox "アクティブシートをCSV形式で保存しました"
End Sub

Sub シート上の画像に枠線を追加する()
    Const MAX_IMAGE_COUNT As Long = 100

    Dim imageCount As Long
    Dim continueProcessing As Boolean

    ' 画像の数をカウント
    imageCount = ActiveSheet.Shapes.count

    ' 画像の数が一定数を超えている場合、警告を表示して処理続行の確認を取得
    If imageCount > MAX_IMAGE_COUNT Then
        continueProcessing = MsgBox("画像の数が " & imageCount & " 個あります。処理時間が長くなる可能性がありますが、続行しますか?", vbQuestion + vbYesNo) = vbYes
    Else
        continueProcessing = True
    End If

    ' 処理を続行する場合、画像に枠線を追加
    If continueProcessing Then
        Dim shape As shape
        
        For Each shape In ActiveSheet.Shapes
            ' 画像のみに枠線を追加
            If shape.Type = msoPicture Then
                shape.Line.Weight = 0.5   ' 枠線の太さを設定
                shape.Line.ForeColor.RGB = RGB(0, 0, 0)  ' 枠線の色を黒に設定
            End If
        Next shape

        Application.StatusBar = "画像に枠線を追加しました。"

        ' ステータスバーの表示を5秒後に消す
        Application.OnTime Now + TimeValue("00:00:05"), "CommonModule.ResetStatusBar"
    Else
        MsgBox "処理を中止しました。", vbInformation
    End If
End Sub

Sub ブック上の全シートの画像に枠線を追加する()
    Const MAX_IMAGE_COUNT As Long = 100

    Dim imageCount As Long
    Dim continueProcessing As Boolean
    
    Dim shape As shape
    Dim sh As Worksheet

    For Each sh In Worksheets
        For Each shape In sh.Shapes
            ' 画像のみに枠線を追加
            If shape.Type = msoPicture Then
                shape.Line.Weight = 0.5   ' 枠線の太さを設定
                shape.Line.ForeColor.RGB = RGB(0, 0, 0) ' 枠線の色を黒に設定
            End If
        Next shape
    Next sh

    Application.StatusBar = "画像に枠線を追加しました。"

    ' ステータスバーの表示を5秒後に消す
    Application.OnTime Now + TimeValue("00:00:05"), "CommonModule.ResetStatusBar"
End Sub


Sub 選択中のシートの列幅を揃える()
    ' 選択中のシートを取得
    Dim selectedSheet As Worksheet
    Set selectedSheet = ActiveSheet

    Dim lastColumn As Long
    lastColumn = selectedSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Dim sh As Worksheet
    Dim i As Long
    For Each sh In ActiveWindow.SelectedSheets
        If Not selectedSheet Is sh Then
            ' 各シートの列幅の揃え
            For i = 1 To lastColumn
                sh.Columns(i).ColumnWidth = selectedSheet.Columns(i).ColumnWidth
            Next i
        End If
    Next sh

    ' ステータスバーへの出力とクリア
    Application.StatusBar = "処理が完了しました。"
    Application.DisplayStatusBar = True

    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Sub 選択中のシートの列幅とフォントを揃える()
    ' 選択中のシートを取得
    Dim selectedSheet As Worksheet
    Set selectedSheet = ActiveSheet

    Dim lastColumn As Long
    lastColumn = selectedSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    ' フォントの揃え
    Dim activeFontName As String
    On Error Resume Next
    activeFontName = selectedSheet.Cells(1, 1).Font.Name
    On Error GoTo 0

    Dim sh As Worksheet
    Dim i As Long
    For Each sh In ActiveWindow.SelectedSheets
        If Not selectedSheet Is sh Then
            ' 各シートの列幅の揃え
            For i = 1 To lastColumn
                sh.Columns(i).ColumnWidth = selectedSheet.Columns(i).ColumnWidth
            Next i
            
            If activeFontName <> "" Then
                sh.Cells.Font.Name = activeFontName
            End If
        End If
    Next sh

    ' ステータスバーへの出力とクリア
    Application.StatusBar = "処理が完了しました。"
    Application.DisplayStatusBar = True

    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

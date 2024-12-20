Attribute VB_Name = "SearchModule"
Option Explicit

' --------------------------------------------------------------------------
' 以下の参照設定が必要です。
' 設定は、[ツール]→[参照設定]で。
' <正規表現検索>
' "Microsoft VBScript Regular Expressions 5.5"
' --------------------------------------------------------------------------

Private Const COLOR_DEFAULT = -1

' --------------------------------------------------------------------------
' strPattern            検索文字列 (正規表現で指定)
' sh                    対象のワークシート
' ignoreCase            大文字と小文字を区別する場合はFalse、区別しない場合はTrue
' color_sheet           検索ヒットしたときにシートの色を変更する場合は色を指定。変更しない場合は-1を指定。
' color_cell            検索ヒットしたときにセルの色の塗りつぶしを変更する場合は色を指定。変更しない場合は-1を指定。
' color_font            検索ヒットしたときに該当箇所のフォントの色を変更する場合は色を指定。変更しない場合は-1を指定。
' regexSearch           正規表現で検索するかどうか。Trueの場合は正規表現で検索する
' blnCellColorFlg       検索ヒットした箇所のセルを塗りつぶすか
' blnFontColorFlg       検索ヒットした箇所の文字色を変更するか
' blnFontColorResetFlg  検索ヒットした箇所の文字色を変更する場合、一回セル内の文字を黒でリセットしてから色をつけるか(true: リセットする, false: リセットしない)
' boldflag              検索ヒットした箇所を太字にするかしないか。Trueの場合太字にする
' underlineflag         検索ヒットした箇所に下線を引くか引かないか。Trueの場合は下線を引く
' strikethroughflag     検索ヒットした箇所に取り消し線を引くか引かないか。Trueの場合は下線を引く
' markTopFlag           検索ヒットした箇所の列の上に★を付けるか付けないか (true:付ける, false:付けない)
' jumpFirstHitCell      検索ヒットしたときに､最初にヒットしたセルにジャンプさせるかとうか｡Trueの場合はジャンブする｡
' targetedSelectedCell  検索対象の範囲を、選択したセルに限定する場合、Trueをセットする
' formatAndFormula      書式、数式を反映して検索するか (true:反映したうえで検索, false:反映しないうえで検索)。デフォルトはfalse。
' --------------------------------------------------------------------------
Public Function func検索ヒットした部分を強調表示(ByVal strPattern As String, _
                        ByVal sh As Worksheet, _
                        ByVal ignoreCase As Boolean, _
                        ByVal color_sheet As Long, _
                        ByVal color_cell As Long, _
                        ByVal color_font As Long, _
                        Optional regexSearch As Boolean = True, _
                        Optional blnCellColorFlg As Boolean = True, _
                        Optional blnFontColorFlg As Boolean = True, _
                        Optional blnFontColorResetFlg As Boolean = False, _
                        Optional boldflag As Boolean = False, _
                        Optional underlineflag As Boolean = False, _
                        Optional strikethroughflag As Boolean = False, _
                        Optional markTopFlag As Boolean = False, _
                        Optional jumpFirstHitCell As Boolean = False, _
                        Optional targetedSelectedCell As Boolean = False, _
                        Optional formatAndFormula As Boolean = False _
                    ) As String

    Dim reg As New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim startPos As Long
    Dim topRow As Long ' 先頭に★マークをつけるように、選択セルの一番上のY座標を記憶する変数
    Dim iLen
    Dim r As Range
    Dim iPosition
    Dim i
    Dim count As Long: count = 0
    Dim resultMessage As String: resultMessage = ""
    Dim targetRange As Range
    Dim cellStr As String
    
    ' 検索文字列
    iLen = Len(strPattern)
    If iLen = 0 Then
        Exit Function
    End If
    
    If markTopFlag Then
        topRow = Rows.count
        For Each r In Selection
            If r.Row < topRow Then
                topRow = r.Row
                If r.Row < topRow Then
                    If topRow = 1 Then
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    
    If targetedSelectedCell Then
        Set targetRange = Selection
    Else
        Set targetRange = sh.UsedRange
    End If
    
    ' 正規表現の条件設定。
    reg.Global = True ' 文字列の最後まで検索(True:する、False:しない)
    reg.ignoreCase = ignoreCase ' 大文字と小文字を区別する場合はFalse、区別しない場合はTrue
    reg.Pattern = strPattern

    ' シートの色をクリアー
    If color_sheet <> COLOR_DEFAULT Then
        sh.Tab.colorIndex = xlNone
    End If

    count = 0
    
    ' 範囲を1セルずつループ
    For Each r In targetRange
        If Not IsError(r.Value) Then
            If r.Value <> vbNullString Then
                ' 正規表現で検索する場合
                If regexSearch Then
                    iPosition = 0
                    
                    ' セル文字列から正規表現での検索を行う
                    If formatAndFormula Then
                        cellStr = r.Value
                    Else
                        cellStr = r.Text
                    End If
                    Set oMatches = reg.Execute(cellStr)
                    
                    ' 検索で見つかった箇所の数をループ
                    For i = 0 To oMatches.count - 1
                        ' 見つかった場合、シートの色を変更
                        If color_sheet <> COLOR_DEFAULT Then
                            sh.Tab.color = color_sheet
                        End If
    
                        ' 見つかった個数をカウント
                        count = count + 1
                        
                        ' 見つかった箇所を取得
                        Set oMatch = oMatches.Item(i)
                        
                        ' 検索一致の先頭位置を取得
                        iPosition = oMatch.FirstIndex
    
                        ' 検索一致文字列長を取得
                        iLen = oMatch.length
                        
                        If i = 0 Then
                            If blnCellColorFlg Then
                                ' セルの塗りつぶし
                                r.Interior.color = color_cell
                            End If
                            
                            If blnFontColorResetFlg Then
                                ' 数式を含むセルの場合は文字色の変更はしない
                                If Not r.HasFormula Then
                                    ' セル内の文字色をいちど黒にする
                                    r.Font.color = 0
                                End If
                            End If
                            
                            ' 検索ヒットしたセルに移動
                            If jumpFirstHitCell Then
                                If count = 1 Then
                                    r.Activate
                                End If
                            End If
                        End If
                        
                        ' 数式を含まないセルの場合は文字色や太字等の変更もする
                        If Not r.HasFormula Then
                            ' 検索一致部分のフォントを変更
                            If boldflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Bold = True   ' 太字
                            End If
                            If underlineflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Underline = True    ' 下線を引く
                            End If
                            If strikethroughflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Strikethrough = True    ' 取り消し線を引く
                            End If
                            If blnFontColorFlg Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.color = color_font    ' フォント色
                            End If
                        End If
                        
                        ' 先頭セルに★マークを付ける場合
                        If (markTopFlag And (topRow > 1)) Then
                            If Cells(topRow - 1, r.Column).Value = "" Then
                                Cells(topRow - 1, r.Column).Value = "★"
                            End If
                        End If
                    Next
                ' 通常の検索をする場合(正規表現ではない場合)
                Else
                    startPos = 1    ' 何文字目から検索するか
                    iPosition = -1  ' 何文字目でヒットしたか。初期値はとりあえず-1で。
                    i = 0           ' このセル内でいくつ見つかったか
                    Do
                        If formatAndFormula Then
                            cellStr = r.Value
                        Else
                            cellStr = r.Text
                        End If
                    
                        ' 大文字小文字を区別しない場合
                        If ignoreCase Then
                            ' テキストモードで比較する(大文字・小文字を区別しない、半角・全角を区別しない)
                            iPosition = InStr(startPos, cellStr, strPattern, vbTextCompare)
                        ' 大文字小文字を区別する場合
                        Else
                            ' バイナリモードで比較する(大文字・小文字を区別する、半角・全角を区別する)
                            iPosition = InStr(startPos, cellStr, strPattern, vbBinaryCompare)
                        End If
                        
                        ' 見つかった場合
                        If iPosition > 0 Then
                            ' 見つかった個数をカウント
                            count = count + 1
                            i = i + 1
                            
                            ' このシート内で初めてヒットした場合
                            If count = 1 Then
                                ' 見つかった場合、シートの色を変更
                                If color_sheet <> COLOR_DEFAULT Then
                                    sh.Tab.color = color_sheet
                                End If
                                
                                ' 検索ヒットしたセルに移動
                                If jumpFirstHitCell Then
                                    r.Activate
                                End If
                            End If
                            
                            ' このセル内で初めてヒットした場合
                            If i = 1 Then
                                If blnCellColorFlg Then
                                    ' セルの塗りつぶし
                                    r.Interior.color = color_cell
                                End If
                                
                                ' 数式を含むセルの場合は文字色の変更はしない
                                If Not r.HasFormula Then
                                    If blnFontColorResetFlg Then
                                        ' セル内の文字色をいちど黒にする
                                        r.Font.color = 0
                                    End If
                                End If
                            End If
                            
                            ' 数式を含まないセルの場合は文字色や太字等の変更もする
                            If Not r.HasFormula Then
                                ' 検索一致部分のフォントを変更
                                If boldflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Bold = True   ' 太字
                                End If
                                If underlineflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Underline = True    ' 下線
                                End If
                                If strikethroughflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Strikethrough = True ' 取り消し線
                                End If
                                If blnFontColorFlg Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.color = color_font     ' フォント色
                                End If
                            End If
                            
                            ' 先頭セルに★マークを付ける場合
                            If (markTopFlag And (topRow > 1)) Then
                                If Cells(topRow - 1, r.Column).Value = "" Then
                                    Cells(topRow - 1, r.Column).Value = "★"
                                End If
                            End If
                            
                            startPos = iPosition + iLen
                        End If
                    Loop While iPosition <> 0
                End If
            End If
        End If
    Next
    
    If count <> 0 Then
        func検索ヒットした部分を強調表示 = sh.Name & ":" & count & "件, "
    End If

End Function

' TODO: 以下Functionは未使用？なら削除してもいいかも。
' --------------------------------------------------------------------------
' a_sht                 ワークシート
' a_sPattern            検索パターン
' a_bIgnoreCase         大文字小文字の区別（True：区別しない、False：区別する）
' a_bFindReplace = True 検索と置換のどちらか（True：検索、False：置換）
' a_sReplace = ""       置換文字列
' --------------------------------------------------------------------------
Function funcFindCellRegExp(a_sht As Worksheet, a_sPattern As String, a_bIgnoreCase As Boolean, Optional a_bFindReplace As Boolean = True, Optional a_sReplace As String = "") As Range
    Dim reg         As New RegExp       '// 正規表現クラス
    Dim iLen                            '// 検索一致文字列長
    Dim r           As Range            '// 選択セル範囲の処理中の１セル
    Dim i                               '// ループカウンタ
    Dim bResult     As Boolean          '// 検索結果
    Dim rPre        As Range            '// アクティブセルより上のセルで一致したセル
    Dim rFind       As Range            '// 検索一致セル
    
    '// 検索文字列が未設定の場合
    iLen = Len(a_sPattern)
    If iLen = 0 Then
        Set funcFindCellRegExp = Nothing
        Exit Function
    End If
    
    '// 正規表現の条件設定
    reg.Global = True               '// 文字列の最後まで検索（True：する、False：しない）
    reg.ignoreCase = a_bIgnoreCase  '// 大文字小文字の区別（True：する、False：しない）
    reg.Pattern = a_sPattern        '// 検索する正規表現パターン
    
    '// セル範囲を１セルずつループ
    For Each r In a_sht.UsedRange
        '// セル文字列から正規表現での検索を行う
        bResult = reg.Test(r.Value)
        
        '// 検索に一致しなかった場合
        If bResult = False Then
            GoTo CONTINUE
        End If
        
        '// 以下検索に一致した場合
        
        Debug.Print r.Address(False, False)
        
        '// 上セルでの検索一致で見つかったセルがまだ無い場合
        If rPre Is Nothing Then
            '// 現在見つかっているセルを設定
            Set rPre = Range(r.Address)
        End If
        
        '// ループのセルがアクティブセルより上にある場合
        If (r.Row < ActiveCell.Row) Then
            GoTo CONTINUE
        '// ループのセルがアクティブセルと同じ行だけど右にある場合
        ElseIf (r.Row = ActiveCell.Row) And (r.Column <= ActiveCell.Column) Then
            GoTo CONTINUE
        '// ループのセルがアクティブセルより右下にある場合
        Else
            '// 検索一致セルが未設定の場合
            If rFind Is Nothing Then
                Set rFind = Range(r.Address)
            End If
        End If
        
CONTINUE:
    Next
    
    '// 見つかった場合
    If Not rFind Is Nothing Then
        Set funcFindCellRegExp = rFind
        'rFind.Select
    '// アクティブセルより上側で見つかった場合
    ElseIf Not rPre Is Nothing Then
        Set funcFindCellRegExp = rPre
        'rPre.Select
    '// 見つからなかった場合
    Else
        Set funcFindCellRegExp = Nothing
        'Call MsgBox("検索対象が見つかりません", vbExclamation, "正規表現検索")
        Exit Function
    End If
    
    '// 置換の場合
    If a_bFindReplace = False Then
        '// アクティブセルの文字列を置換
        ActiveCell.Value = reg.Replace(ActiveCell.Value, a_sReplace)
        Set funcFindCellRegExp = ActiveCell
    End If
End Function

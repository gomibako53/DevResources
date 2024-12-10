Attribute VB_Name = "CommonModule"
Option Explicit

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

' クリップボード関数
Public Function PutClipBoard(str As String)
    Dim temp As String

    ' クリップボードに文字列を格納
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = str
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

End Function

' -----------------------------------------
' 正規表現でパターンにマッチするか判定
' strTest       判定用の文字列
' strPattern    正規表現パターン
' ignoreCase    大文字と小文字を区別しない(デフォルト：True)
' -----------------------------------------
Public Function regularExpressionTest(strTest As String, strPattern As String, Optional ignoreCase As Boolean = True)
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")
    regularExpressionTest = False

    With RE
        .Pattern = strPattern       ' 検索パターンを設定
        .ignoreCase = ignoreCase    ' 大文字と小文字を区別しない
        .Global = True              ' 文字列全体を検索
        If .Test(strTest) Then
            regularExpressionTest = True
        End If
    End With
    Set RE = Nothing
End Function

' -----------------------------------------
' 指定範囲内で文字列の置換
' rng           指定範囲
' strTarget     検索文字列
' strReplace    置換文字列
' -----------------------------------------
Public Function replaceSpecifiedRange(rng As Range, strTarget As String, strReplace As String)
    Dim TgRng As Range
    Dim FRng As Range
    Dim Fst As String

    Set TgRng = rng
    Set FRng = rng.Find(strTarget, LookAt:=xlPart)
    If Not FRng Is Nothing Then
        Fst = FRng.Address
        Do
            FRng.Value = Replace(FRng.Text, strTarget, strReplace)
            Set FRng = TgRng.FindNext(FRng)
        Loop Until FRng Is Nothing
    End If
    Set FRng = Nothing
    Set TgRng = Nothing
End Function

' -----------------------------------------
' ステータスバーの表示をクリアする
' -----------------------------------------
Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub


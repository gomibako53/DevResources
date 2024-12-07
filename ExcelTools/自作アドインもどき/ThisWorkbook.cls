VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =====================
' ワークブックオープン
' =====================
Public Sub Workbook_Opne()
    ' マクロのショートカットキー設定
    SetShortcutKeys
End Sub

' =====================
' ショートカットキー設定
' =====================
Public Sub SetShortcutKeys()
    ' [Ctrl+Shift+g]    : 特定キーワードの強調表示
    Application.OnKey "+^g", "ThisWorkbook.特定キーワードの強調表示"
End Sub

Public Sub 特定キーワードの強調表示()
    ' シート検索用ユーザフォームを表示
    frmSearchAndHighlightWords.Show vbModeless
    ' 入力テキストボックスをフォーカス
    frmSearchAndHighlightWords.InputVal.SetFocus
End Sub

Public Sub 正規表現セル検索()
    ' シート検索用ユーザフォームを表示
    frmRegexSearch.Show vbModeless
    ' 入力テキストボックスをフォーカス
    frmRegexSearch.InputVal.SetFocus
End Sub


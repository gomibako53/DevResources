VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegexSearch 
   Caption         =   "正規表現セル検索"
   ClientHeight    =   1845
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "frmRegexSearch.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRegexSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_INPUT_VAL = ""
Private Const WORK_FILE = "C:\kan\tools\MacroWorkFiles\frmRegexSearch.txt"


' ----------------------------------------------
' 選択したセルのみ検索対象チェックボックス処理
' ----------------------------------------------
Private Sub SelectedCellCheckBox_Click()
    If SelectedCellCheckBox Then
        SearchAreaCheckBox = False
    End If

End Sub

' ----------------------------------------------
' 全シート対象チェックボックス処理
' ----------------------------------------------
Private Sub SearchAreaCheckBox_Click()
    If SearchAreaCheckBox Then
        SelectedCellCheckBox = False
    End If
End Sub

' ----------------------------------------------
' 検索ボタン押下処理
' ----------------------------------------------
Private Sub SearchButton_Click()
' --------------------------------------------------------------------------
' a_sht                 ワークシート
' a_sPattern            検索パターン
' a_bIgnoreCase         大文字小文字の区別（True：区別しない、False：区別する）
' a_bFindReplace = True 検索と置換のどちらか（True：検索、False：置換）
' a_sReplace = ""       置換文字列
' --------------------------------------------------------------------------
'Function FindCellRegExp(a_sht As Worksheet, a_sPattern As String, a_bIgnoreCase As Boolean, Optional a_bFindReplace As Boolean = True, Optional a_sReplace As String = "") As Range
    Dim sh As Worksheet
    Dim rng As Range
    
    Set sh = ActiveSheet
    Set rng = funcFindCellRegExp(sh, InputVal, CaseSensitiveCheckBox)
    
    ' ヒットしたものが無い場合
    If rng Is Nothing Then
        Exit Sub
    End If

    rng.Select
End Sub

' ----------------------------------------------
' 閉じる
' ----------------------------------------------
Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Function ReadWorkFile_FirstLine() As String
    Dim buf As String
    
    ' ファイルが存在しない場合
    If Dir(WORK_FILE) = "" Then
        ReadWorkFile_FirstLine = ""
        End Function
    End If

    Open WORK_FILE For Input As #1
        Line Input #1, buf
    Close #1

    If Not IsEmpty(buf) Then
        ReadWorkFile_FirstLine = buf
        workFile_lastFirstLine = buf    ' AppendWorkFile_FirstLineのために、1行目の内容を記憶しておく
    End If
End Function

Private Sub AppendWorkFile_FirstLine(ByVal str As String)
    Dim fso As Object
    Dim buf As String
    
    ' 変更がない場合(テキストファイルの1行目と、今回の内容が変更ない場合)は、処理しない
    If workFile_lastFirstLine = str Then
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    buf = fso.OpenTextFile(WORK_FILE).ReadAll
    
    buf = Replace(buf, vbCrLf & str & vbCrLf, vbCrLf)
    buf = Replace(buf, vbCrLf & vbCrLf, vbCrLf)
    buf = Replace(buf, vbCrLf & vbCrLf, vbCrLf)
    buf = str & vbCrLf & buf
    
    Open WORK_FILE For Output As #1
        Print #1, buf
    Close #1
End Sub



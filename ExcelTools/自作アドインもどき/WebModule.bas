Attribute VB_Name = "WebModule"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                       (ByVal hwnd As Long, ByVal lpOperation As String, _
                        ByVal lpFile As String, ByVal lpParameters As String, _
                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub 先頭範囲のセルに書かれたURLを開く()
'    Dim top As Long
'    Dim bottom As Long
'    Dim left As Long
'    Dim right As Long
    Dim c As Range
    Dim url As String
    Dim rc
    
'    ' マクロ実行時の選択範囲の座標を取得
'    top = Selection(1).row
'    left = Selection(1).Column
'    bottom = Selection(Selection.Count).row
'    right = Selection(Selection.Count).Column
'    'Debug.Print selectionTop & "'" & selectionLeft & "'" & selectionBottom & "'" & selectionRight
    
    ' 先頭セルを選択
    For Each c In Selection
        ' 非表示のセルでなければ
        If c.EntireRow.Hidden = False Then
            ' 何か記載されているセルなら
            If c.Text <> "" Then
                url = c.Text
                rc = ShellExecute(0, "Open", url, "", "", 1)
            End If
        End If
    Next c
   

'    URL = "http://www.officetanaka.net/"
'    rc = ShellExecute(0, "Open", URL, "", "", 1)
End Sub



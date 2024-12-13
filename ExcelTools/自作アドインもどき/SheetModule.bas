Attribute VB_Name = "SheetModule"
Option Explicit

Sub シート名を検索して選択()
    Dim sh_find As String
    Dim i As Long
    Dim sheet_no As Long
    Dim msg As String

    Application.StatusBar = False

    sh_find = InputBox("検索するシート名を入力してください。")
    If sh_find = "" Then Exit Sub

    For i = 1 To Sheets.count
        sheet_no = (ActiveSheet.Index + i) Mod Sheets.count
        ' 大文字・小文字を区別しないで検索
        If UCase(Sheets(i).Name) Like UCase("*" & sh_find & "*") Then
            If Sheets(i).Visible = True Then
                Sheets(i).Select
                Exit Sub
            End If
        End If
    Next i

    msg = "指定された文字列をシート名に含むシートは見つかりませんでした。"
    Application.StatusBar = msg
End Sub

Attribute VB_Name = "SheetModule"
Option Explicit

Sub �V�[�g�����������đI��()
    Dim sh_find As String
    Dim i As Long
    Dim sheet_no As Long
    Dim msg As String

    Application.StatusBar = False

    sh_find = InputBox("��������V�[�g������͂��Ă��������B")
    If sh_find = "" Then Exit Sub

    For i = 1 To Sheets.count
        sheet_no = (ActiveSheet.Index + i) Mod Sheets.count
        ' �啶���E����������ʂ��Ȃ��Ō���
        If UCase(Sheets(i).Name) Like UCase("*" & sh_find & "*") Then
            If Sheets(i).Visible = True Then
                Sheets(i).Select
                Exit Sub
            End If
        End If
    Next i

    msg = "�w�肳�ꂽ��������V�[�g���Ɋ܂ރV�[�g�͌�����܂���ł����B"
    Application.StatusBar = msg
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchAndHighlightWords 
   Caption         =   "検索する文字列を入力してください"
   ClientHeight    =   4330
   ClientLeft      =   -45
   ClientTop       =   285
   ClientWidth     =   5310
   OleObjectBlob   =   "frmSearchAndHighlightWords.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSearchAndHighlightWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_INPUT_VAL = ""
Private Const WORK_FILE = "C:\kan\tools\MacroWorkFiles\frmSearchAndHighlightWords.txt"
Private Const INI_FILE = "C:\kan\tools\MacroWorkFiles\frmSearchAndHighlightWords.ini"

Private workFile_lastFirstLine As String

Private blnSearchAreaFlg As Boolean ' 全シートをターゲットにするかどうか(true:する, false:しない)
Private blnSelectedCellFlg As Boolean ' 選択したセルのみを検索対象とするか(true:する, false:しない)
Private blnRegexFlg As Boolean ' 正規表現で検索するか(true:する, false:しない)
Private blnCellColorFlg As Boolean ' 検索ヒットした箇所のセルを塗りつぶすか(true:する, false:しない)
Private blnFontColorFlg As Boolean ' 検索ヒットした箇所の文字色を変更するか(true:する, false:しない)
Private blnBoldFlg As Boolean ' 検索ヒットした箇所を太字にするかしないか(true:する, false:しない)

Private color_sheet As Long
Private color_cell As Long
Private color_font As Long

Private Const COLOR_DEFAULT = -1
Private Const COLOR_BLACK = &H0
Private Const COLOR_WHITE = &HFFFFFF
Private Const COLOR_RED = &HFF
Private Const COLOR_BLUE = &HFF0000
Private Const COLOR_GREEN = &HFF00
Private Const COLOR_YELLOW = &HFFFF&
Private Const COLOR_PINK = &HFF00FF
Private Const COLOR_ORANGE = &H80FF&

Private Sub ClearBackColorButton_Click()
    color_cell = COLOR_DEFAULT
    CellColorSampleLabel.BackColor = COLOR_WHITE
End Sub

Private Sub ClearFontColorButton_Click()
    color_font = COLOR_DEFAULT
    CellColorSampleLabel.ForeColor = COLOR_BLACK
End Sub

Private Sub ClearSheetColorButton_Click()
    color_sheet = COLOR_DEFAULT
    SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
End Sub

Private Sub UserForm_Initialize()
    ' デフォルト色を設定
    color_cell = COLOR_YELLOW
    color_font = COLOR_RED
    color_sheet = COLOR_ORANGE
    
    CellColorSampleLabel.BackColor = color_cell
    CellColorSampleLabel.ForeColor = color_font
    
    ' 全シートをターゲットにする場合
    If blnSearchAreaFlg Then
        SheetColorSampleLabel.BackColor = color_sheet
    ' 単一シートのみをターゲットにする場合
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
    End If

    InputVal.Text = ReadWorkFile_FirstLine
End Sub

' ----------------------------------------------
' 検索ボタン押下処理
' ----------------------------------------------
Private Sub SearchButton_Click()

    Dim sh As Worksheet
    Dim resultMessage As String: resultMessage = ""
    Dim strTemp As String: strTemp = ""
    
    Dim strPattern As String: strPattern = ""
    
    If SheetColorSampleLabel.BackColor = &H8000000F Then
        color_sheet = COLOR_DEFAULT
    Else
        color_sheet = SheetColorSampleLabel.BackColor
    End If
    
    ' TODO
    If SheetColorSampleLabel.BackColor = COLOR_WHITE Then
        color_cell = COLOR_DEFAULT
    Else
        color_cell = CellColorSampleLabel.BackColor
    End If
    
    ' TODO
    If SheetColorSampleLabel.BackColor = COLOR_BLACK Then
        color_font = COLOR_DEFAULT
    Else
        color_font = CellColorSampleLabel.ForeColor
    End If
    
    ' 入力値取得(テキスト)
    strPattern = Replace(InputVal.Text, vbTab, "")

    ' 入力値が未入力の場合、処理を中断する
    If strPattern = vbNullString Then Exit Sub

    strTemp = "検索実行中です。 ： """ & strPattern & """"
    If Len(strTemp) > 100 Then
        strTemp = left(strTemp, 100)
    End If
    Application.StatusBar = "検索実行中です。 ：""" & strTemp & """"
    
    ' 入力値取得(検索条件)
    blnSearchAreaFlg = SearchAreaCheckBox ' 全シート対象にするかどうか
    blnSelectedCellFlg = SelectedCellCheckBox ' 選択したセルのみを検索対象とするか
    blnRegexFlg = RegexCheckBox ' 正規表現で検索するか
    blnCellColorFlg = CellColorCheckBox  ' 検索ヒットした箇所のセルを塗りつぶすか
    blnFontColorFlg = FontColorCheckBox ' 検索ヒットした箇所の文字色を変更するか
    blnBoldFlg = BoldCheckBox ' 検索ヒットした箇所を太字にするかしないか
    
    Application.ScreenUpdating = False '処理中は画面描画をOFFにする場合はココのコメントを外す
    
    Call AppendWorkFile_FirstLine(strPattern)
    
    Unload Me
    
    ' 全シートをターゲットにする場合
    If blnSearchAreaFlg Then
        For Each sh In Worksheets
            strTemp = func検索ヒットした部分を強調表示(strPattern, sh, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnBoldFlg)
            If strTemp <> "" Then
                resultMessage = resultMessage & strTemp
            End If
        Next sh
    ' 単ーシートのみをターゲットにする場合
    Else
        color_sheet = COLOR_DEFAULT ' 検索ヒットしてもシートの色は変更しない
        strTemp = func検索ヒットした部分を強調表示(strPattern, ActiveSheet, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnBoldFlg, True, blnSelectedCellFlg)
        If strTemp <> "" Then
            resultMessage = resultMessage & strTemp
        End If
    End If
    
    If resultMessage <> "" Then
        If Len(resultMessage) > 100 Then
            resultMessage = left(resultMessage, 100)
        End If
        Application.StatusBar = resultMessage
    Else
        Application.StatusBar = "検索ヒットしたセルはありませんでした。：""" & strPattern & """"
    End If
    
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------
' 閉じる
' ----------------------------------------------
Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub SearchAreaCheckBox_Click()
    blnSearchAreaFlg = SearchAreaCheckBox   ' 全シート対象にするかどうか

    ' 全シートをターゲットにする場合
    If blnSearchAreaFlg Then
        SelectedCellCheckBox = False

        If color_sheet = -1 Then
            color_sheet = COLOR_RED
        End If
        SheetColorSampleLabel.BackColor = color_sheet
    ' 単-シートのみをターゲットにする場合
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
    End If
End Sub

Private Sub BackColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(color_cell)

    If chosenColor <> COLOR_DEFAULT Then
        CellColorSampleLabel.BackColor = chosenColor
    End If
End Sub

Private Sub FontColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(color_font)

    If chosenColor <> COLOR_DEFAULT Then
        CellColorSampleLabel.ForeColor = chosenColor
    End If
End Sub

Private Sub SheetColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(color_sheet)

    If chosenColor <> COLOR_DEFAULT Then
        SheetColorSampleLabel.BackColor = chosenColor
    End If
End Sub

Private Function ReadWorkFile_FirstLine() As String
    Dim buf As String

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


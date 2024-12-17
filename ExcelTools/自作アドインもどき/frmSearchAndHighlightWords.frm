VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchAndHighlightWords 
   Caption         =   "検索する文字列を入力してください"
   ClientHeight    =   5970
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
Private Const WORK_FILE_NAME = "frmSearchAndHighlightWords.txt"
Private WORK_FILE_PATH As String
Private Const INI_FILE_NAME = "frmSearchAndHighlightWords.ini"
Private INI_FILE_PATH As String

Private workFile_lastFirstLine As String

Private blnSearchAreaFlg As Boolean ' 全シートをターゲットにするかどうか(true:する, false:しない)
Private blnSelectedCellFlg As Boolean ' 選択したセルのみを検索対象とするか(true:する, false:しない)
Private blnRegexFlg As Boolean ' 正規表現で検索するか(true:する, false:しない)
Private blnCellColorFlg As Boolean ' 検索ヒットした箇所のセルを塗りつぶすか(true:する, false:しない)
Private blnFontColorFlg As Boolean ' 検索ヒットした箇所の文字色を変更するか(true:する, false:しない)
Private blnFontColorResetFlg As Boolean ' 検索ヒットした箇所の文字色を変更する場合、一回セル内の文字を黒でリセットしてから色をつけるか (true:リセットする、false:リセットしない)
Private blnBoldFlg As Boolean ' 検索ヒットした箇所を太字にするかしないか(true:する, false:しない)
Private blnUnderlineFlg As Boolean ' 検索ヒットした箇所に下線を引くか引かないか (true:引く, false:引かない)
Private blnStrikethroughFlg As Boolean ' 検索ヒットした箇所に取り消し線を引くか引かないか (true:引く, false:引かない)
Private blnMarkTopFlg As Boolean ' 検索ヒットした箇所の列の上に★を付けるか付けないか(true:付ける, false:付けない)
Private blnFormatFormulaFlg As Boolean ' 書式,数式を反映して検索するか (true:反映したうえで検索, false:反映しないうえで検索)

Private Const COLOR_DEFAULT = -1
Private Const COLOR_BLACK = &H0
Private Const COLOR_WHITE = &HFFFFFF
Private Const COLOR_RED = &HFF
Private Const COLOR_BLUE = &HFF0000
Private Const COLOR_GREEN = &HFF00
Private Const COLOR_YELLOW = &HFFFF&
Private Const COLOR_PINK = &HFF00FF
Private Const COLOR_ORANGE = &H80FF&
Private Const COLOR_LIGHT_YELLOW = &HCCFFFF
Private Const COLOR_LIGHT_PINK = &HFFCCFF
Private Const COLOR_LIGHT_BLUE = &HFFFFCC
Private Const COLOR_LIGHT_GREEN = &HCCFFCC

Private Sub UserForm_Initialize()
    ' フォーム表示位置の初期化。StartUpPosition=2だけにすると、マルチウインドウでうまく動かない場合があるので。
    frmSearchAndHighlightWords.StartUpPosition = 2
    frmSearchAndHighlightWords.Top = Application.Top + ((Application.Height - frmSearchAndHighlightWords.Height) / 2)
    frmSearchAndHighlightWords.Left = Application.Left + ((Application.Width - frmSearchAndHighlightWords.Width) / 2)

    WORK_FILE_PATH = ThisWorkbook.Path & "\" & WORK_FILE_NAME
    ' ファイルが存在しない場合は新規作成
    If Dir(WORK_FILE_PATH) = "" Then
        Open WORK_FILE_PATH For Output As #1
        Close #1
    End If
    INI_FILE_PATH = ThisWorkbook.Path & "\" & INI_FILE_NAME
    If Dir(INI_FILE_PATH) = "" Then
        Open INI_FILE_PATH For Output As #1
        Close #1
    End If

    ' チェックボックスのON/OFFのデフォルト値
    Call ReadIniFile

    ' 単一シートのみをターゲットにする場合
    If Not blnSearchAreaFlg Then
        SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
    End If

    InputVal.Text = ReadWorkFile_FirstLine
End Sub

Private Sub ClearBackColorButton_Click()
    CellColorSampleLabel.BackColor = COLOR_WHITE
End Sub

Private Sub ClearFontColorButton_Click()
    CellColorSampleLabel.ForeColor = COLOR_BLACK
End Sub

Private Sub ClearSheetColorButton_Click()
    SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
End Sub

' ----------------------------------------------
' 検索ボタン押下処理
' ----------------------------------------------
Private Sub SearchButton_Click()

    Dim sh As Worksheet
    Dim resultMessage As String: resultMessage = ""
    Dim strTemp As String: strTemp = ""
    
    Dim strPattern As String: strPattern = ""
    
    Dim color_sheet As Long
    Dim color_cell As Long
    Dim color_font As Long
    
    If SheetColorSampleLabel.BackColor = &H8000000F Then
        color_sheet = COLOR_DEFAULT
    Else
        color_sheet = SheetColorSampleLabel.BackColor
    End If
    
    If SheetColorSampleLabel.BackColor = COLOR_WHITE Then
        color_cell = COLOR_DEFAULT
    Else
        color_cell = CellColorSampleLabel.BackColor
    End If
    
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
        strTemp = Left(strTemp, 100)
    End If
    Application.StatusBar = "検索実行中です。 ：""" & strTemp & """"
    
    ' 検索条件のファイル保存
    Call WriteIniFile
    
    ' 入力値取得(検索条件)
    blnSearchAreaFlg = SearchAreaCheckBox ' 全シート対象にするかどうか
    blnSelectedCellFlg = SelectedCellCheckBox ' 選択したセルのみを検索対象とするか
    blnRegexFlg = RegexCheckBox ' 正規表現で検索するか
    blnCellColorFlg = CellColorCheckBox  ' 検索ヒットした箇所のセルを塗りつぶすか
    blnFontColorFlg = FontColorCheckBox ' 検索ヒットした箇所の文字色を変更するか
    blnFontColorResetFlg = FontColorResetCheckBox   ' セル内の文字を黒でリセットしてから色をつけるか
    blnBoldFlg = BoldCheckBox ' 検索ヒットした箇所を太字にするかしないか
    blnUnderlineFlg = UnderlineCheckBox ' 検索ヒットした箇所に下線を引くか引かないか
    blnStrikethroughFlg = StrikethroughCheckBox ' 検索ヒットした箇所に取り消し線を引くか引かないか
    blnMarkTopFlg = MarkTopCheckBox ' 検索ヒットした箇所の列の上に★を付けるか付けないか
    If blnSelectedCellFlg = False Then
        blnMarkTopFlg = False ' シート内全範囲が対象の場合は★マークは強制的に付けない
    End If
    blnFormatFormulaFlg = FormatFormulaCheckBox ' 書式,数式を反映して検索するか
    
    Application.ScreenUpdating = False '処理中は画面描画をOFFにする場合はココのコメントを外す
    
    Call AppendWorkFile_FirstLine(strPattern)
    
    Unload Me
    
    ' 全シートをターゲットにする場合
    If blnSearchAreaFlg Then
        For Each sh In Worksheets
            strTemp = func検索ヒットした部分を強調表示(strPattern, sh, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnFontColorResetFlg, blnBoldFlg, blnUnderlineFlg, blnStrikethroughFlg, blnMarkTopFlg, False, False, blnFormatFormulaFlg)
            If strTemp <> "" Then
                resultMessage = resultMessage & strTemp
            End If
        Next sh
    ' 単ーシートのみをターゲットにする場合
    Else
        color_sheet = COLOR_DEFAULT ' 検索ヒットしてもシートの色は変更しない
        strTemp = func検索ヒットした部分を強調表示(strPattern, ActiveSheet, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnFontColorResetFlg, blnBoldFlg, blnUnderlineFlg, blnStrikethroughFlg, blnMarkTopFlg, True, blnSelectedCellFlg, blnFormatFormulaFlg)
        If strTemp <> "" Then
            resultMessage = resultMessage & strTemp
        End If
    End If
    
    If resultMessage <> "" Then
        ' 結果をクリップボードに貼り付け。どのシートに何件ヒットしたのか。
        Dim clipboardMessage As String: clipboardMessage = resultMessage
        clipboardMessage = Left(clipboardMessage, Len(clipboardMessage) - 2) ' 終端の「, 」を除去
        clipboardMessage = Replace(clipboardMessage, "件, ", "件" & vbCrLf) ' 「件, 」を「件<改行>」に置換
        CommonModule.PutClipBoard clipboardMessage
    
        If Len(resultMessage) > 100 Then
            resultMessage = Left(resultMessage, 100)
        End If
        Application.StatusBar = resultMessage
    Else
        strTemp = "検索ヒットしたセルはありませんでした。：""" & strPattern & """"
        If LenB(strTemp) > 497 Then
            strTemp = LeftB(strTemp, 496)
        End If
        Application.StatusBar = strTemp
    End If
    
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------
' 閉じる
' ----------------------------------------------
Private Sub CloseButton_Click()
    Call WriteIniFile
    Unload Me
End Sub

Private Sub SelectedCellCheckBox_Click()
    ' シート全体が検索対象の場合
    If Not SelectedCellCheckBox Then
        MarkTopCheckBox.value = False
    End If
End Sub

Private Sub SearchAreaCheckBox_Click()
    blnSearchAreaFlg = SearchAreaCheckBox   ' 全シート対象にするかどうか

    ' 全シートをターゲットにする場合
    If blnSearchAreaFlg Then
        MarkTopCheckBox.value = False
        
        SelectedCellCheckBox = False

        SheetColorSampleLabel.BackColor = COLOR_ORANGE
    ' 単-シートのみをターゲットにする場合
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' 検索ヒットしてもシートの色は変更しない
    End If
End Sub

Private Sub BackColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(CellColorSampleLabel.BackColor)

    If chosenColor <> COLOR_DEFAULT Then
        CellColorSampleLabel.BackColor = chosenColor
    End If
End Sub

Private Sub FontColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(CellColorSampleLabel.ForeColor)

    If chosenColor <> COLOR_DEFAULT Then
        CellColorSampleLabel.ForeColor = chosenColor
    End If
End Sub

Private Sub SheetColorButton_Click()
    Dim chosenColor As Long
    chosenColor = GetColorDlg(SheetColorSampleLabel.BackColor)

    If chosenColor <> COLOR_DEFAULT Then
        SheetColorSampleLabel.BackColor = chosenColor
    End If
End Sub

Private Function ReadWorkFile_FirstLine() As String
    Dim buf As String
    Dim fileLength As Long

    Open WORK_FILE_PATH For Input As #1
        fileLength = LOF(1)

        If fileLength > 0 Then
            Line Input #1, buf
        Else
            buf = ""
        End If
    Close #1

    If Not IsEmpty(buf) Then
        ReadWorkFile_FirstLine = buf
        workFile_lastFirstLine = buf    ' AppendWorkFile_FirstLineのために、1行目の内容を記憶しておく
    Else
        ReadWorkFile_FirstLine = ""
    End If
End Function

Private Sub AppendWorkFile_FirstLine(ByVal str As String)
    Dim FSO As Object
    Dim buf As String
    
    ' 変更がない場合(テキストファイルの1行目と、今回の内容が変更ない場合)は、処理しない
    If workFile_lastFirstLine = str Then
        Exit Sub
    End If
    
    ' 書き込み対象が空文字を指定された場合は処理を抜ける。このIF文は絶対ひっかからないハズだけど安全のため。
    If IsEmpty(str) Or str = "" Then
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If IsEmpty(workFile_lastFirstLine) Or workFile_lastFirstLine = "" Then
        ' 空ファイルに初めて新規で追加する場合
        buf = str
    Else
        buf = FSO.OpenTextFile(WORK_FILE_PATH).ReadAll
    
        buf = Replace(buf, vbCrLf & str & vbCrLf, vbCrLf)
        buf = Replace(buf, vbCrLf & vbCrLf, vbCrLf)
        buf = Replace(buf, vbCrLf & vbCrLf, vbCrLf)
        buf = str & vbCrLf & buf
    End If
    
    Open WORK_FILE_PATH For Output As #1
        Print #1, buf
    Close #1
End Sub

' INIファイルの読み込み関数
Private Sub ReadIniFile()
    Dim FSO As Object
    Dim iniFile As Object
    Dim line As String
    Dim key As String
    Dim value As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set iniFile = FSO.OpenTextFile(INI_FILE_PATH, 1)

    ' デフォルト値 (INIファイルに定義がない場合のデフォルト設定)
    RegexCheckBox = True ' 正規表現で検索するか
    CellColorCheckBox = True ' 検索ヒットした箇所のセルを塗りつぶすか
    FontColorCheckBox = True ' 検索ヒットした箇所の文字色を変更するか
    BoldCheckBox = False ' 検索ヒットした箇所を太字にするかしないか
    UnderlineCheckBox = False ' 検索ヒットした箇所に下線を引くか引かないか
    StrikethroughCheckBox = False ' 検索ヒットした箇所に取り消し線を引くか引かないか
    MarkTopCheckBox = False ' 検索ヒットした箇所の列の上に★を付けるか付けないか
    SelectedCellCheckBox = True ' 選択したセルのみを検索対象とするか
    SearchAreaCheckBox = False  ' 全シート対象にするかどうか
    FormatFormulaCheckBox = False ' 書式,数式を反映して検索するか
    ' デフォルト色を設定
    CellColorSampleLabel.BackColor = COLOR_LIGHT_YELLOW
    CellColorSampleLabel.ForeColor = COLOR_RED
    SheetColorSampleLabel.BackColor = COLOR_ORANGE

    Do While Not iniFile.AtEndOfStream
        line = iniFile.ReadLine
        If InStr(line, "=") > 0 Then
            key = Trim(Left(line, InStr(line, "=") - 1))
            value = Trim(Mid(line, InStr(line, "=") + 1))
            Select Case key
                Case "RegexCheckBox"
                    RegexCheckBox = CBool(value)
                Case "CellColorCheckBox"
                    CellColorCheckBox = CBool(value)
                Case "FontColorCheckBox"
                    FontColorCheckBox = CBool(value)
                Case "BoldCheckBox"
                    BoldCheckBox = CBool(value)
                Case "UnderlineCheckBox"
                    UnderlineCheckBox = CBool(value)
                Case "StrikethroughCheckBox"
                    StrikethroughCheckBox = CBool(value)
                Case "MarkTopCheckBox"
                    MarkTopCheckBox = CBool(value)
                Case "SelectedCellCheckBox"
                    SelectedCellCheckBox = CBool(value)
                Case "SearchAreaCheckBox"
                    SearchAreaCheckBox = CBool(value)
                Case "FormatFormulaCheckBox"
                    FormatFormulaCheckBox = CBool(value)
                Case "ColorCell"
                    CellColorSampleLabel.BackColor = CLng(value)
                Case "ColorFont"
                    CellColorSampleLabel.ForeColor = CLng(value)
                Case "ColorSheet"
                    SheetColorSampleLabel.BackColor = CLng(value)
            End Select
        End If
    Loop

    iniFile.Close

    ' ---------------------------------------------
    ' 一部の設定はINIファイルを無視して上書き
    ' ---------------------------------------------

    If SelectedCellCheckBox = False Then
        MarkTopCheckBox = False ' シート内全範囲が対象の場合は★マークは強制的に付けない
    End If

    ' 単一シートのみをターゲットにする場合
    If Not SearchAreaCheckBox Then
        SheetColorSampleLabel.BackColor = &H8000000F  ' 検索ヒットしてもシートの色は変更しない
    End If
End Sub

' INIファイルの書き込み関数
Private Sub WriteIniFile()
    Dim FSO As Object
    Dim iniFile As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set iniFile = FSO.CreateTextFile(INI_FILE_PATH, True)

    iniFile.WriteLine "RegexCheckBox=" & CStr(RegexCheckBox)
    iniFile.WriteLine "CellColorCheckBox=" & CStr(CellColorCheckBox)
    iniFile.WriteLine "FontColorCheckBox=" & CStr(FontColorCheckBox)
    iniFile.WriteLine "BoldCheckBox=" & CStr(BoldCheckBox)
    iniFile.WriteLine "UnderlineCheckBox=" & CStr(UnderlineCheckBox)
    iniFile.WriteLine "StrikethroughCheckBox=" & CStr(StrikethroughCheckBox)
    iniFile.WriteLine "MarkTopCheckBox=" & CStr(MarkTopCheckBox)
    iniFile.WriteLine "SelectedCellCheckBox=" & CStr(SelectedCellCheckBox)
    iniFile.WriteLine "SearchAreaCheckBox=" & CStr(SearchAreaCheckBox)
    iniFile.WriteLine "FormatFormulaCheckBox=" & CStr(FormatFormulaCheckBox)
    iniFile.WriteLine "ColorCell=" & CStr(CellColorSampleLabel.BackColor)
    iniFile.WriteLine "ColorFont=" & CStr(CellColorSampleLabel.ForeColor)
    iniFile.WriteLine "ColorSheet=" & CStr(SheetColorSampleLabel.BackColor)

    iniFile.Close
End Sub

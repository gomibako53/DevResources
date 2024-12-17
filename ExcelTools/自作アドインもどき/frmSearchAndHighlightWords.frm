VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchAndHighlightWords 
   Caption         =   "�������镶�������͂��Ă�������"
   ClientHeight    =   5970
   ClientLeft      =   -45
   ClientTop       =   285
   ClientWidth     =   5310
   OleObjectBlob   =   "frmSearchAndHighlightWords.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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

Private blnSearchAreaFlg As Boolean ' �S�V�[�g���^�[�Q�b�g�ɂ��邩�ǂ���(true:����, false:���Ȃ�)
Private blnSelectedCellFlg As Boolean ' �I�������Z���݂̂������ΏۂƂ��邩(true:����, false:���Ȃ�)
Private blnRegexFlg As Boolean ' ���K�\���Ō������邩(true:����, false:���Ȃ�)
Private blnCellColorFlg As Boolean ' �����q�b�g�����ӏ��̃Z����h��Ԃ���(true:����, false:���Ȃ�)
Private blnFontColorFlg As Boolean ' �����q�b�g�����ӏ��̕����F��ύX���邩(true:����, false:���Ȃ�)
Private blnFontColorResetFlg As Boolean ' �����q�b�g�����ӏ��̕����F��ύX����ꍇ�A���Z�����̕��������Ń��Z�b�g���Ă���F�����邩 (true:���Z�b�g����Afalse:���Z�b�g���Ȃ�)
Private blnBoldFlg As Boolean ' �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ���(true:����, false:���Ȃ�)
Private blnUnderlineFlg As Boolean ' �����q�b�g�����ӏ��ɉ����������������Ȃ��� (true:����, false:�����Ȃ�)
Private blnStrikethroughFlg As Boolean ' �����q�b�g�����ӏ��Ɏ��������������������Ȃ��� (true:����, false:�����Ȃ�)
Private blnMarkTopFlg As Boolean ' �����q�b�g�����ӏ��̗�̏�Ɂ���t���邩�t���Ȃ���(true:�t����, false:�t���Ȃ�)
Private blnFormatFormulaFlg As Boolean ' ����,�����𔽉f���Č������邩 (true:���f���������Ō���, false:���f���Ȃ������Ō���)

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
    ' �t�H�[���\���ʒu�̏������BStartUpPosition=2�����ɂ���ƁA�}���`�E�C���h�E�ł��܂������Ȃ��ꍇ������̂ŁB
    frmSearchAndHighlightWords.StartUpPosition = 2
    frmSearchAndHighlightWords.Top = Application.Top + ((Application.Height - frmSearchAndHighlightWords.Height) / 2)
    frmSearchAndHighlightWords.Left = Application.Left + ((Application.Width - frmSearchAndHighlightWords.Width) / 2)

    WORK_FILE_PATH = ThisWorkbook.Path & "\" & WORK_FILE_NAME
    ' �t�@�C�������݂��Ȃ��ꍇ�͐V�K�쐬
    If Dir(WORK_FILE_PATH) = "" Then
        Open WORK_FILE_PATH For Output As #1
        Close #1
    End If
    INI_FILE_PATH = ThisWorkbook.Path & "\" & INI_FILE_NAME
    If Dir(INI_FILE_PATH) = "" Then
        Open INI_FILE_PATH For Output As #1
        Close #1
    End If

    ' �`�F�b�N�{�b�N�X��ON/OFF�̃f�t�H���g�l
    Call ReadIniFile

    ' �P��V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    If Not blnSearchAreaFlg Then
        SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
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
    SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
End Sub

' ----------------------------------------------
' �����{�^����������
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
    
    ' ���͒l�擾(�e�L�X�g)
    strPattern = Replace(InputVal.Text, vbTab, "")

    ' ���͒l�������͂̏ꍇ�A�����𒆒f����
    If strPattern = vbNullString Then Exit Sub

    strTemp = "�������s���ł��B �F """ & strPattern & """"
    If Len(strTemp) > 100 Then
        strTemp = Left(strTemp, 100)
    End If
    Application.StatusBar = "�������s���ł��B �F""" & strTemp & """"
    
    ' ���������̃t�@�C���ۑ�
    Call WriteIniFile
    
    ' ���͒l�擾(��������)
    blnSearchAreaFlg = SearchAreaCheckBox ' �S�V�[�g�Ώۂɂ��邩�ǂ���
    blnSelectedCellFlg = SelectedCellCheckBox ' �I�������Z���݂̂������ΏۂƂ��邩
    blnRegexFlg = RegexCheckBox ' ���K�\���Ō������邩
    blnCellColorFlg = CellColorCheckBox  ' �����q�b�g�����ӏ��̃Z����h��Ԃ���
    blnFontColorFlg = FontColorCheckBox ' �����q�b�g�����ӏ��̕����F��ύX���邩
    blnFontColorResetFlg = FontColorResetCheckBox   ' �Z�����̕��������Ń��Z�b�g���Ă���F�����邩
    blnBoldFlg = BoldCheckBox ' �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ���
    blnUnderlineFlg = UnderlineCheckBox ' �����q�b�g�����ӏ��ɉ����������������Ȃ���
    blnStrikethroughFlg = StrikethroughCheckBox ' �����q�b�g�����ӏ��Ɏ��������������������Ȃ���
    blnMarkTopFlg = MarkTopCheckBox ' �����q�b�g�����ӏ��̗�̏�Ɂ���t���邩�t���Ȃ���
    If blnSelectedCellFlg = False Then
        blnMarkTopFlg = False ' �V�[�g���S�͈͂��Ώۂ̏ꍇ�́��}�[�N�͋����I�ɕt���Ȃ�
    End If
    blnFormatFormulaFlg = FormatFormulaCheckBox ' ����,�����𔽉f���Č������邩
    
    Application.ScreenUpdating = False '�������͉�ʕ`���OFF�ɂ���ꍇ�̓R�R�̃R�����g���O��
    
    Call AppendWorkFile_FirstLine(strPattern)
    
    Unload Me
    
    ' �S�V�[�g���^�[�Q�b�g�ɂ���ꍇ
    If blnSearchAreaFlg Then
        For Each sh In Worksheets
            strTemp = func�����q�b�g���������������\��(strPattern, sh, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnFontColorResetFlg, blnBoldFlg, blnUnderlineFlg, blnStrikethroughFlg, blnMarkTopFlg, False, False, blnFormatFormulaFlg)
            If strTemp <> "" Then
                resultMessage = resultMessage & strTemp
            End If
        Next sh
    ' �P�[�V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    Else
        color_sheet = COLOR_DEFAULT ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
        strTemp = func�����q�b�g���������������\��(strPattern, ActiveSheet, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnFontColorResetFlg, blnBoldFlg, blnUnderlineFlg, blnStrikethroughFlg, blnMarkTopFlg, True, blnSelectedCellFlg, blnFormatFormulaFlg)
        If strTemp <> "" Then
            resultMessage = resultMessage & strTemp
        End If
    End If
    
    If resultMessage <> "" Then
        ' ���ʂ��N���b�v�{�[�h�ɓ\��t���B�ǂ̃V�[�g�ɉ����q�b�g�����̂��B
        Dim clipboardMessage As String: clipboardMessage = resultMessage
        clipboardMessage = Left(clipboardMessage, Len(clipboardMessage) - 2) ' �I�[�́u, �v������
        clipboardMessage = Replace(clipboardMessage, "��, ", "��" & vbCrLf) ' �u��, �v���u��<���s>�v�ɒu��
        CommonModule.PutClipBoard clipboardMessage
    
        If Len(resultMessage) > 100 Then
            resultMessage = Left(resultMessage, 100)
        End If
        Application.StatusBar = resultMessage
    Else
        strTemp = "�����q�b�g�����Z���͂���܂���ł����B�F""" & strPattern & """"
        If LenB(strTemp) > 497 Then
            strTemp = LeftB(strTemp, 496)
        End If
        Application.StatusBar = strTemp
    End If
    
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------
' ����
' ----------------------------------------------
Private Sub CloseButton_Click()
    Call WriteIniFile
    Unload Me
End Sub

Private Sub SelectedCellCheckBox_Click()
    ' �V�[�g�S�̂������Ώۂ̏ꍇ
    If Not SelectedCellCheckBox Then
        MarkTopCheckBox.value = False
    End If
End Sub

Private Sub SearchAreaCheckBox_Click()
    blnSearchAreaFlg = SearchAreaCheckBox   ' �S�V�[�g�Ώۂɂ��邩�ǂ���

    ' �S�V�[�g���^�[�Q�b�g�ɂ���ꍇ
    If blnSearchAreaFlg Then
        MarkTopCheckBox.value = False
        
        SelectedCellCheckBox = False

        SheetColorSampleLabel.BackColor = COLOR_ORANGE
    ' �P-�V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
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
        workFile_lastFirstLine = buf    ' AppendWorkFile_FirstLine�̂��߂ɁA1�s�ڂ̓��e���L�����Ă���
    Else
        ReadWorkFile_FirstLine = ""
    End If
End Function

Private Sub AppendWorkFile_FirstLine(ByVal str As String)
    Dim FSO As Object
    Dim buf As String
    
    ' �ύX���Ȃ��ꍇ(�e�L�X�g�t�@�C����1�s�ڂƁA����̓��e���ύX�Ȃ��ꍇ)�́A�������Ȃ�
    If workFile_lastFirstLine = str Then
        Exit Sub
    End If
    
    ' �������ݑΏۂ��󕶎����w�肳�ꂽ�ꍇ�͏����𔲂���B����IF���͐�΂Ђ�������Ȃ��n�Y�����ǈ��S�̂��߁B
    If IsEmpty(str) Or str = "" Then
        Exit Sub
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If IsEmpty(workFile_lastFirstLine) Or workFile_lastFirstLine = "" Then
        ' ��t�@�C���ɏ��߂ĐV�K�Œǉ�����ꍇ
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

' INI�t�@�C���̓ǂݍ��݊֐�
Private Sub ReadIniFile()
    Dim FSO As Object
    Dim iniFile As Object
    Dim line As String
    Dim key As String
    Dim value As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set iniFile = FSO.OpenTextFile(INI_FILE_PATH, 1)

    ' �f�t�H���g�l (INI�t�@�C���ɒ�`���Ȃ��ꍇ�̃f�t�H���g�ݒ�)
    RegexCheckBox = True ' ���K�\���Ō������邩
    CellColorCheckBox = True ' �����q�b�g�����ӏ��̃Z����h��Ԃ���
    FontColorCheckBox = True ' �����q�b�g�����ӏ��̕����F��ύX���邩
    BoldCheckBox = False ' �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ���
    UnderlineCheckBox = False ' �����q�b�g�����ӏ��ɉ����������������Ȃ���
    StrikethroughCheckBox = False ' �����q�b�g�����ӏ��Ɏ��������������������Ȃ���
    MarkTopCheckBox = False ' �����q�b�g�����ӏ��̗�̏�Ɂ���t���邩�t���Ȃ���
    SelectedCellCheckBox = True ' �I�������Z���݂̂������ΏۂƂ��邩
    SearchAreaCheckBox = False  ' �S�V�[�g�Ώۂɂ��邩�ǂ���
    FormatFormulaCheckBox = False ' ����,�����𔽉f���Č������邩
    ' �f�t�H���g�F��ݒ�
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
    ' �ꕔ�̐ݒ��INI�t�@�C���𖳎����ď㏑��
    ' ---------------------------------------------

    If SelectedCellCheckBox = False Then
        MarkTopCheckBox = False ' �V�[�g���S�͈͂��Ώۂ̏ꍇ�́��}�[�N�͋����I�ɕt���Ȃ�
    End If

    ' �P��V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    If Not SearchAreaCheckBox Then
        SheetColorSampleLabel.BackColor = &H8000000F  ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
    End If
End Sub

' INI�t�@�C���̏������݊֐�
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

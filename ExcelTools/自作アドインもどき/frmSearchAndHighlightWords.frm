VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchAndHighlightWords 
   Caption         =   "�������镶�������͂��Ă�������"
   ClientHeight    =   4330
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
Private Const WORK_FILE = "C:\kan\tools\MacroWorkFiles\frmSearchAndHighlightWords.txt"
Private Const INI_FILE = "C:\kan\tools\MacroWorkFiles\frmSearchAndHighlightWords.ini"

Private workFile_lastFirstLine As String

Private blnSearchAreaFlg As Boolean ' �S�V�[�g���^�[�Q�b�g�ɂ��邩�ǂ���(true:����, false:���Ȃ�)
Private blnSelectedCellFlg As Boolean ' �I�������Z���݂̂������ΏۂƂ��邩(true:����, false:���Ȃ�)
Private blnRegexFlg As Boolean ' ���K�\���Ō������邩(true:����, false:���Ȃ�)
Private blnCellColorFlg As Boolean ' �����q�b�g�����ӏ��̃Z����h��Ԃ���(true:����, false:���Ȃ�)
Private blnFontColorFlg As Boolean ' �����q�b�g�����ӏ��̕����F��ύX���邩(true:����, false:���Ȃ�)
Private blnBoldFlg As Boolean ' �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ���(true:����, false:���Ȃ�)

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
    SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
End Sub

Private Sub UserForm_Initialize()
    ' �f�t�H���g�F��ݒ�
    color_cell = COLOR_YELLOW
    color_font = COLOR_RED
    color_sheet = COLOR_ORANGE
    
    CellColorSampleLabel.BackColor = color_cell
    CellColorSampleLabel.ForeColor = color_font
    
    ' �S�V�[�g���^�[�Q�b�g�ɂ���ꍇ
    If blnSearchAreaFlg Then
        SheetColorSampleLabel.BackColor = color_sheet
    ' �P��V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
    End If

    InputVal.Text = ReadWorkFile_FirstLine
End Sub

' ----------------------------------------------
' �����{�^����������
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
    
    ' ���͒l�擾(�e�L�X�g)
    strPattern = Replace(InputVal.Text, vbTab, "")

    ' ���͒l�������͂̏ꍇ�A�����𒆒f����
    If strPattern = vbNullString Then Exit Sub

    strTemp = "�������s���ł��B �F """ & strPattern & """"
    If Len(strTemp) > 100 Then
        strTemp = left(strTemp, 100)
    End If
    Application.StatusBar = "�������s���ł��B �F""" & strTemp & """"
    
    ' ���͒l�擾(��������)
    blnSearchAreaFlg = SearchAreaCheckBox ' �S�V�[�g�Ώۂɂ��邩�ǂ���
    blnSelectedCellFlg = SelectedCellCheckBox ' �I�������Z���݂̂������ΏۂƂ��邩
    blnRegexFlg = RegexCheckBox ' ���K�\���Ō������邩
    blnCellColorFlg = CellColorCheckBox  ' �����q�b�g�����ӏ��̃Z����h��Ԃ���
    blnFontColorFlg = FontColorCheckBox ' �����q�b�g�����ӏ��̕����F��ύX���邩
    blnBoldFlg = BoldCheckBox ' �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ���
    
    Application.ScreenUpdating = False '�������͉�ʕ`���OFF�ɂ���ꍇ�̓R�R�̃R�����g���O��
    
    Call AppendWorkFile_FirstLine(strPattern)
    
    Unload Me
    
    ' �S�V�[�g���^�[�Q�b�g�ɂ���ꍇ
    If blnSearchAreaFlg Then
        For Each sh In Worksheets
            strTemp = func�����q�b�g���������������\��(strPattern, sh, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnBoldFlg)
            If strTemp <> "" Then
                resultMessage = resultMessage & strTemp
            End If
        Next sh
    ' �P�[�V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    Else
        color_sheet = COLOR_DEFAULT ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
        strTemp = func�����q�b�g���������������\��(strPattern, ActiveSheet, True, color_sheet, color_cell, color_font, blnRegexFlg, blnCellColorFlg, blnFontColorFlg, blnBoldFlg, True, blnSelectedCellFlg)
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
        Application.StatusBar = "�����q�b�g�����Z���͂���܂���ł����B�F""" & strPattern & """"
    End If
    
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------
' ����
' ----------------------------------------------
Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub SearchAreaCheckBox_Click()
    blnSearchAreaFlg = SearchAreaCheckBox   ' �S�V�[�g�Ώۂɂ��邩�ǂ���

    ' �S�V�[�g���^�[�Q�b�g�ɂ���ꍇ
    If blnSearchAreaFlg Then
        SelectedCellCheckBox = False

        If color_sheet = -1 Then
            color_sheet = COLOR_RED
        End If
        SheetColorSampleLabel.BackColor = color_sheet
    ' �P-�V�[�g�݂̂��^�[�Q�b�g�ɂ���ꍇ
    Else
        SheetColorSampleLabel.BackColor = &H8000000F    ' �����q�b�g���Ă��V�[�g�̐F�͕ύX���Ȃ�
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
        workFile_lastFirstLine = buf    ' AppendWorkFile_FirstLine�̂��߂ɁA1�s�ڂ̓��e���L�����Ă���
    End If
End Function

Private Sub AppendWorkFile_FirstLine(ByVal str As String)
    Dim fso As Object
    Dim buf As String
    
    ' �ύX���Ȃ��ꍇ(�e�L�X�g�t�@�C����1�s�ڂƁA����̓��e���ύX�Ȃ��ꍇ)�́A�������Ȃ�
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


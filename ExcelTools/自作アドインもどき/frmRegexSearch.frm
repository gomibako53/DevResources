VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegexSearch 
   Caption         =   "���K�\���Z������"
   ClientHeight    =   1845
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "frmRegexSearch.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
' �I�������Z���̂݌����Ώۃ`�F�b�N�{�b�N�X����
' ----------------------------------------------
Private Sub SelectedCellCheckBox_Click()
    If SelectedCellCheckBox Then
        SearchAreaCheckBox = False
    End If

End Sub

' ----------------------------------------------
' �S�V�[�g�Ώۃ`�F�b�N�{�b�N�X����
' ----------------------------------------------
Private Sub SearchAreaCheckBox_Click()
    If SearchAreaCheckBox Then
        SelectedCellCheckBox = False
    End If
End Sub

' ----------------------------------------------
' �����{�^����������
' ----------------------------------------------
Private Sub SearchButton_Click()
' --------------------------------------------------------------------------
' a_sht                 ���[�N�V�[�g
' a_sPattern            �����p�^�[��
' a_bIgnoreCase         �啶���������̋�ʁiTrue�F��ʂ��Ȃ��AFalse�F��ʂ���j
' a_bFindReplace = True �����ƒu���̂ǂ��炩�iTrue�F�����AFalse�F�u���j
' a_sReplace = ""       �u��������
' --------------------------------------------------------------------------
'Function FindCellRegExp(a_sht As Worksheet, a_sPattern As String, a_bIgnoreCase As Boolean, Optional a_bFindReplace As Boolean = True, Optional a_sReplace As String = "") As Range
    Dim sh As Worksheet
    Dim rng As Range
    
    Set sh = ActiveSheet
    Set rng = funcFindCellRegExp(sh, InputVal, CaseSensitiveCheckBox)
    
    ' �q�b�g�������̂������ꍇ
    If rng Is Nothing Then
        Exit Sub
    End If

    rng.Select
End Sub

' ----------------------------------------------
' ����
' ----------------------------------------------
Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Function ReadWorkFile_FirstLine() As String
    Dim buf As String
    
    ' �t�@�C�������݂��Ȃ��ꍇ
    If Dir(WORK_FILE) = "" Then
        ReadWorkFile_FirstLine = ""
        End Function
    End If

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



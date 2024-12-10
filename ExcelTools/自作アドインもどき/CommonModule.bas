Attribute VB_Name = "CommonModule"
Option Explicit

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

' �N���b�v�{�[�h�֐�
Public Function PutClipBoard(str As String)
    Dim temp As String

    ' �N���b�v�{�[�h�ɕ�������i�[
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = str
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

End Function

' -----------------------------------------
' ���K�\���Ńp�^�[���Ƀ}�b�`���邩����
' strTest       ����p�̕�����
' strPattern    ���K�\���p�^�[��
' ignoreCase    �啶���Ə���������ʂ��Ȃ�(�f�t�H���g�FTrue)
' -----------------------------------------
Public Function regularExpressionTest(strTest As String, strPattern As String, Optional ignoreCase As Boolean = True)
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")
    regularExpressionTest = False

    With RE
        .Pattern = strPattern       ' �����p�^�[����ݒ�
        .ignoreCase = ignoreCase    ' �啶���Ə���������ʂ��Ȃ�
        .Global = True              ' ������S�̂�����
        If .Test(strTest) Then
            regularExpressionTest = True
        End If
    End With
    Set RE = Nothing
End Function

' -----------------------------------------
' �w��͈͓��ŕ�����̒u��
' rng           �w��͈�
' strTarget     ����������
' strReplace    �u��������
' -----------------------------------------
Public Function replaceSpecifiedRange(rng As Range, strTarget As String, strReplace As String)
    Dim TgRng As Range
    Dim FRng As Range
    Dim Fst As String

    Set TgRng = rng
    Set FRng = rng.Find(strTarget, LookAt:=xlPart)
    If Not FRng Is Nothing Then
        Fst = FRng.Address
        Do
            FRng.Value = Replace(FRng.Text, strTarget, strReplace)
            Set FRng = TgRng.FindNext(FRng)
        Loop Until FRng Is Nothing
    End If
    Set FRng = Nothing
    Set TgRng = Nothing
End Function

' -----------------------------------------
' �X�e�[�^�X�o�[�̕\�����N���A����
' -----------------------------------------
Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub


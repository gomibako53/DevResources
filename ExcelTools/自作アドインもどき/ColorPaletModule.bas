Attribute VB_Name = "ColorPaletModule"
Option Explicit

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChooseColor As ChooseColor) As Long

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type



Private Const CC_RGBINT = &H1 ' �F�̃f�t�H���g�l��ݒ�
Private Const CC_LFULLOPEN = &H2 ' �F�̍쐬���s��������\��
Private Const CC_PREVENTFULLOPEN = &H4 ' �F�̍쐬�{�^���𖳌��ɂ���
Private Const CC_SHOWHELP = &H8 ' �w���v�{�^����\��

' �@�\  �F  �F�I���_�C�A���O��\�����A�I�����ꂽ�F��RGB�l��Ԃ�
' ����  �F  lngDefColor �f�t�H���g�\������F
' �Ԓl  �F  ������ RGB�l�A�L�����Z���� -1�A�G���[�� -2 (�[���͍��Ȃ̂Œ���)
Public Function GetColorDlg(lngDefColor As Long) As Long
    Dim udtChooseColor As ChooseColor
    Dim lngRet As Long

    With udtChooseColor
        ' �_�C�A���O�̐ݒ�
        .lStructSize = Len(udtChooseColor)
        .IpCustColors = String$(64, Chr$(0))
        '.flags CC_RGBINT + CC_LFULLOPEN
        .flags = 0
        .rgbResult = lngDefColor
        ' �_�C�A���O�̕\��
        lngRet = ChooseColor(udtChooseColor)
        ' �_�C�A���O����̕Ԃ�l���`�F�b�N
        If lngRet <> 0 Then
            If .rgbResult > RGB(255, 255, 255) Then
                ' �G���[
                GetColorDlg = -2
            Else
                ' ����I���ARGB�l��Ԃ�l�ɃZ�b�g
                GetColorDlg = .rgbResult
            End If
        Else
            ' �L�����Z���������ꂽ�Ƃ�
            GetColorDlg = -1
        End If
    End With
End Function


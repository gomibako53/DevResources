Attribute VB_Name = "ColorPaletModule"
Option Explicit

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChooseColor As ChooseColor) As Long

Private Const CC_RGBINIT = &H1 ' �F�̃f�t�H���g�l��ݒ�
Private Const CC_FULLOPEN = &H2 ' �F�̍쐬���s��������\��(�E���̕���)
Private Const CC_PREVENTFULLOPEN = &H4 ' �F�̍쐬�{�^���𖳌��ɂ���
Private Const CC_SHOWHELP = &H8 ' �w���v�{�^����\��
Private Const CC_ANYCOLOR = &H100 ' ���p�\�Ȋ�{�F�����ׂĕ\��

' �@�\  �F  �F�̐ݒ�_�C�A���O��\�����A�����őI�����ꂽ�F��RGB�̒l��Ԃ�
' ����  �F  lngDefColor �f�t�H���g�\������F
' �Ԓl  �F  ������ RGB�l�A�L�����Z���� -1�A�G���[�� -2 (�[���͍��Ȃ̂Œ���)
Public Function GetColorDlg(lngDefColor As Long) As Long
    Dim udtChooseColor As ChooseColor
    Dim lngRet As LongPtr
    Static CustomColors(16) As Long
    
    ' Some predefined color, there are 16 slots available for predefined colors
    CustomColors(0) = RGB(255, 255, 255)    ' White
    CustomColors(1) = RGB(0, 0, 0)  ' Black
    CustomColors(2) = RGB(255, 0, 0)    ' Red
    'CustomColors(3) = RGB(0, 255, 0)   ' Green
    CustomColors(3) = RGB(0, 176, 80)   ' Green(default)
    CustomColors(4) = RGB(0, 0, 255)  ' Blue
    CustomColors(8) = RGB(255, 255, 204)    ' Light Yellow
    CustomColors(9) = RGB(255, 204, 255)  ' Light Pink
    CustomColors(10) = RGB(204, 255, 255)  ' Light Blue
    CustomColors(11) = RGB(204, 255, 204)  ' Light Green
    CustomColors(12) = RGB(191, 191, 191) ' Light Gray

    With udtChooseColor
        .lStructSize = LenB(udtChooseColor)
        .flags = CC_RGBINIT Or CC_ANYCOLOR
        If IsNull(lngDefColor) = False And IsMissing(lngDefColor) = False Then
            .rgbResult = lngDefColor  'Set the initial color of the dialog
        End If
        .lpCustColors = VarPtr(CustomColors(0))
        
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


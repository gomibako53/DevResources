Attribute VB_Name = "ExportModule"
Option Explicit

' --------------------------------------------------------------------------
' �ȉ��̎Q�Ɛݒ肪�K�v�ł��B
' �ݒ�́A[�c�[��]��[�Q�Ɛݒ�]�ŁB
' "Microsoft Visual Basic for Applications Extensibility *.*"
' --------------------------------------------------------------------------
' �ȉ������� VBA ���W���[���̃G�N�X�|�[�g����������܂����A�G�N�Z���̐ݒ��ύX���Ȃ��ƃG���[�ɂȂ�܂��B
' �����G���[�ɂȂ�����ȉ��ݒ���������Ă��������B
'   1. Excel���J���A[�t�@�C��] �^�u���N���b�N���܂��B
'   2. [�I�v�V����] ���N���b�N���܂��B
'   3. [�g���X�g�Z���^�[] ���N���b�N���A[�g���X�g�Z���^�[�̐ݒ�] ���N���b�N���܂��B
'   4. [�}�N���̐ݒ�] ���N���b�N���A[VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������]�̃`�F�b�N�{�b�N�X���I���ɂ��܂��B
'   5. [OK] ���N���b�N���āA�_�C�A���O�{�b�N�X����܂��B
' --------------------------------------------------------------------------

Sub VBA���W���[�����ꊇExport()
    Dim module      As VBComponent      ' ���W���[��
    Dim moduleList  As VBComponents     ' VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                       ' ���W���[���̊g���q
    Dim sPath                           ' �����Ώۃu�b�N�̃p�X
    Dim sFilePath                       ' �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook                      ' �����Ώۃu�b�N�I�u�W�F�N�g

    Set TargetBook = ActiveWorkbook ' �\�����Ă���u�b�N��ΏۂƂ���

    sPath = TargetBook.Path

    ' �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents

    ' VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        ' �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ' �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
            ' .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        ' �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        ' ���̑�
        Else
            ' �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If

        ' �V�����t�H���_�̃p�X���쐬 (�u�b�N������g���q����菜�������̂̃t�H���_)
        Dim newFolderPath As String: newFolderPath = sPath & "\" & "VBA_Export_" & Split(TargetBook.Name, ".")(0)
        ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
        If Dir(newFolderPath, vbDirectory) = "" Then
            MkDir newFolderPath
        End If

        ' �G�N�X�|�[�g���{
        sFilePath = newFolderPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)

        ' �o�͐�m�F�p���O�o��
        Debug.Print sFilePath

CONTINUE:
    Next
End Sub


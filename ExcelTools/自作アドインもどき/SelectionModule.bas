Attribute VB_Name = "SelectionModule"
Option Explicit

' -----------------------------------------------------
' Function �֐�
' -----------------------------------------------------

Private Function func�I��͈͂̍��W���擾(ByRef ref_rangeTop As Long, ByRef ref_rangeBottom As Long, _
                                        ByRef ref_rangeLeft As Long, ByRef ref_rangeRight As Long)
    ref_rangeTop = Selection(1).row
    ref_rangeBottom = Selection(Selection.count).row
    ref_rangeLeft = Selection(1).Column
    ref_rangeRight = Selection(Selection.count).Column
End Function

Private Function func�\�쐬(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional titleLineBackColor As Long = -1, _
                            Optional titleLineLetterColor As Long = xlAutomatic, _
                            Optional titleLineNum As Long = 1)

    ' �S�̂̐��`��
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeRight))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlHairline   ' �������͓_��
    End With
    
    ' �^�C�g���s�̐��`��
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
        .Interior.COLOR = titleLineBackColor
        .Font.FontStyle = "�W��"
        .Font.ColorIndex = titleLineLetterColor
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' �^�C�g���s��2�s�ȏ�̏ꍇ�́A���Ԃ̉����C�����o���Ȃ�
    If titleLineNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
            .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        End With
    End If
End Function

Private Function func_�g�쐬(ByVal COLOR As Double, Optional frameType As Integer = 1)
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    Application.ScreenUpdating = False
    
    With Range(Cells(top, left), Cells(bottom, right))
    
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeLeft).Weight = xlHairline
        .Borders(xlEdgeRight).Weight = xlHairline
        
        If frameType = 1 Then
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        ElseIf frameType = 2 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlHairline   ' �������͓_��
        End If
        
        .Interior.COLOR = COLOR
    End With
    
    Application.ScreenUpdating = True
End Function

' -----------------------------------------------------
' Sub �֐�
' -----------------------------------------------------

Sub �\�쐬_�Z��()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 6299648              ' �Z��
    ' �^�C�g���̕����F
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' ��
    
    Call func�\�쐬(top, bottom, left, right, BackColor, FontColor)
    
    ' ���F�t�H�[�}�b�g�̏ꍇ�̓���ݒ�
    With Range(Cells(top, left), Cells(top, right))
        .Font.FontStyle = "����"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�Z��_2�sver()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 6299648              ' �Z��
    ' �^�C�g���̕����F
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' ��
    
    Call func�\�쐬(top, bottom, left, right, BackColor, FontColor, 2)
    
    ' ���F�t�H�[�}�b�g�̏ꍇ�̓���ݒ�
    With Range(Cells(top, left), Cells(top + 1, right))
        .Font.FontStyle = "����"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    Call func�\�쐬(top, bottom, left, right)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 10092543              ' ���F
    
    Call func�\�쐬(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�I�����W()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 10079487              ' �I�����W
    
    Call func�\�쐬(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_��()
    Application.ScreenUpdating = False
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 13434828              ' ��
    
    Call func�\�쐬(top, bottom, left, right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �Ԃ̏c��������()
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub �g_��()
    func_�g�쐬 (10092543)
End Sub

Sub �g2_��()
    Call func_�g�쐬(10092543, 2)
End Sub

Sub �g_�O���[()
    func_�g�쐬 (12632256)
End Sub

Sub �s�P�ʂŃZ���̌���()
    Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' �s�P�ʂŃZ���̌���
    If (bottom - top) < 5000 Then   ' �I���������Ă���ꍇ�͏������ł܂�\��������̂ŉ������������Ȃ��B
        For i = top To bottom
            Range(Cells(i, left), Cells(i, right)).MergeCells = True
        Next i
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub ��P�ʂŃZ���̌���()
    Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    ' ��P�ʂŃZ���̌���
    For i = left To right
        Range(Cells(top, i), Cells(bottom, i)).MergeCells = True
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub �E�B���h�E�g�̌Œ�����Ȃ���()
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

Sub �I��͈͂̊e�Z����ҏW��Ԃɂ���Enter()
    ' �t�@�C���p�X���L�ڂ��ꂽ�Z������������ҏW��Ԃɂ���Enter�������ƁA�n�C�p�[�����N��������B
    ' ���̊֐��́A�n�C�p�[�����N�������邽�߂Ɏg�p����

    Application.ScreenUpdating = False
    
    Dim i       As Long
    Dim c       As Range
    
    ' �I������Ă���͈͂��擾
    Dim top As Long, bottom As Long, left As Long, right As Long
    Call func�I��͈͂̍��W���擾(top, bottom, left, right)
    
    For Each c In Range(Cells(top, left), Cells(bottom, right))
        If c.Value <> "" Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        ElseIf VarType(c.Value) = vbError Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        End If
    Next c

    Application.ScreenUpdating = True
End Sub



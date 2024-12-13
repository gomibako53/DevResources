Attribute VB_Name = "SelectionModule"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' -----------------------------------------------------
' Function �֐�
' -----------------------------------------------------

Private Function func�I��͈͂̍��W���擾(ByRef ref_rangeTop As Long, ByRef ref_rangeBottom As Long, _
                                        ByRef ref_rangeLeft As Long, ByRef ref_rangeRight As Long)
    ref_rangeTop = Selection(1).Row
    ref_rangeBottom = Selection(Selection.count).Row
    ref_rangeLeft = Selection(1).Column
    ref_rangeRight = Selection(Selection.count).Column
End Function

Private Function func�\�쐬(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional removeVerticalBorders As Boolean = False, _
                            Optional removeHorizontalBorders As Boolean = False, _
                            Optional titleLineBackColor As Long = -1, _
                            Optional titleLineLetterColor As Long = xlAutomatic, _
                            Optional titleLineNum As Long = 1)
    
    Dim i As Long, j As Long

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
        .Interior.color = titleLineBackColor
        .Font.FontStyle = "�W��"
        .Font.colorIndex = titleLineLetterColor
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' �^�C�g���s��2�s�ȏ�̏ꍇ�́A���Ԃ̉����C�����o���Ȃ�
    If titleLineNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeTop + titleLineNum - 1, rangeRight))
            .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        End With
    End If
    
    ' �c�̌r���̃��C�����K�v�Ȃ��ꍇ�͏���
    If removeVerticalBorders Then
        For i = rangeLeft To rangeRight
            Dim emptyCount As Long
            emptyCount = Application.WorksheetFunction.CountIf(Range(Cells(rangeTop, i), Cells(rangeBottom, i)), "<>")
            If emptyCount <= 0 Then
                Range(Cells(rangeTop, i), Cells(rangeBottom, i)).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                Range(Cells(rangeTop, i), Cells(rangeBottom, i)).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
            End If
        Next i
    End If

    Application.ScreenUpdating = True
End Function

Private Function func�\�쐬_��(ByVal rangeTop As Long, ByVal rangeBottom As Long, _
                            ByVal rangeLeft As Long, ByVal rangeRight As Long, _
                            Optional titleColumnBackColor As Long = -1, _
                            Optional titleColumnLetterColor As Long = xlAutomatic, _
                            Optional titleColumnNum As Long = 1)

    ' �S�̂̐��`��
    With Range(Cells(rangeToo, rangeLeft), Cells(rangeBottom, rangeRight))
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
        .Borders(xlInsideHorizontal).Weight = xlHairline    ' �������͓_��
    End With

    ' �^�C�g����̐��`��
    With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeLeft + titleColumnNum - 1))
        .Interior.color = titleColumnBackColor
        .Font.FontStyle = "�W��"
        .Font.colorIndex = titleColumnLetterColor
        .Borders(xlEdgeRight).LineStyle = xlDouble
    End With

    ' �^�C�g����2�s�ȏ�̏ꍇ�́A���Ԃ̏c���C�����o���Ȃ�
    If titleColumnNum >= 2 Then
        With Range(Cells(rangeTop, rangeLeft), Cells(rangeBottom, rangeLeft + titleColumnNum - 1))
            .Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        End With
    End If
    
    Application.ScreenUpdating = True
End Function

Private Function func_�g�쐬(ByVal color As Double, Optional frameType As Integer = 1)
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    With Range(Cells(Top, Left), Cells(bottom, Right))
    
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
            .Borders(xlInsideVertical).LineStyle = xlNone
        ElseIf frameType = 2 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlHairline   ' �������͓_��
            .Borders(xlInsideVertical).LineStyle = xlNone
        End If
        
        .Interior.color = color
    End With
    
    Application.ScreenUpdating = True
End Function

' -----------------------------------------------------
' Sub �֐�
' -----------------------------------------------------

Sub �\�쐬_�Z��()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 6299648              ' �Z��
    ' �^�C�g���̕����F
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' ��
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor, FontColor)
    
    ' ���F�t�H�[�}�b�g�̏ꍇ�̓���ݒ�
    With Range(Cells(Top, Left), Cells(Top, Right))
        .Font.FontStyle = "����"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�Z��_2�sver()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 6299648              ' �Z��
    ' �^�C�g���̕����F
    Dim FontColor As Long: FontColor = xlThemeColorDark1    ' ��
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor, FontColor, 2)
    
    ' ���F�t�H�[�}�b�g�̏ꍇ�̓���ݒ�
    With Range(Cells(Top, Left), Cells(Top + 1, Right))
        .Font.FontStyle = "����"
        .Font.ThemeColor = xlThemeColorDark1
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlEdgeBottom).ThemeColor = 1
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    Call func�\�쐬(Top, bottom, Left, Right)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F_2�s()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, -1, xlAutomatic, 2)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F_��()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    Call func�\�쐬_��(Top, bottom, Left, Right)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_���F()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 10092543              ' ���F
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�I�����W()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 10079487              ' �I�����W
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_��()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 13434828              ' ��
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�O���[()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 15395562              ' ���߂̃O���[
    
    Call func�\�쐬(Top, bottom, Left, Right, True, True, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �\�쐬_�O���[_��()
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' �^�C�g���̔w�i�F
    Dim BackColor As Long: BackColor = 15395562              ' ���߂̃O���[
    
    Call func�\�쐬_��(Top, bottom, Left, Right, BackColor)
    
    Application.ScreenUpdating = True
End Sub

Sub �g_��()
    func_�g�쐬 (10092543)
End Sub

Sub �g2_��()
    Call func_�g�쐬(10092543, 2)
End Sub

Sub �g_�O���[()
    func_�g�쐬 (15395562)  ' ���߂̃O���[
End Sub

Sub ��P�ʂŃZ���̌���()
    'Application.ScreenUpdating = False
    
    Dim i       As Long
    
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    ' ��P�ʂŃZ���̌���
    For i = Left To Right
        Range(Cells(Top, i), Cells(bottom, i)).MergeCells = True
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

    'Application.ScreenUpdating = False
    
    Dim i       As Long
    Dim c       As Range
    
    ' �I������Ă���͈͂��擾
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)
    
    For Each c In Range(Cells(Top, Left), Cells(bottom, Right))
        If c.HasFormula Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        ElseIf c.Value <> "" Then
            SendKeys "{F2}", True
            SendKeys "{ENTER}", True
        End If
        Sleep (500)
    Next c

    Application.ScreenUpdating = True
End Sub

Sub �I�������Z���̃R�����g�ʒu���C��()
    Dim targetRange As Range: Set targetRange = Selection
    Dim myRange As Range

    For Each myRange In targetRange
        If Not (myRange.Comment Is Nothing) Then
            With myRange.Comment.shape
                .Top = myRange.Top
                .Left = myRange.Offset(, 1).Left
                .TextFrame.AutoSize = True
            End With
        End If
    Next
End Sub

Sub �I��͈͂Ńu�����N�̃Z���͏�̃Z���l������()
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Dim i As Long, j As Long

    'Application.ScreenUpdating = False

    ' �I������Ă���͈͂��擾
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)

    For i = Left To Right
        For j = Top To bottom
            ' �}�[�W����Ă���Z���͖���
            If Not Cells(j, i).MergeCells Then
                ' ��\���̃Z���͖���
                If Not Cells(j, i).Rows.Hidden Then
                    If Not Cells(j, i).Columns.Hidden Then
                        ' �l�������Ă��Ȃ������炷����̃Z���̒l�Ɠ������̂�����
                        If j >= 2 And Len(Cells(j, i).Text) = 0 And Len(Cells(j - 1, i).Text) > 0 Then
                            Cells(j, i).Value = Cells(j - 1, i).Value
                        End If
                    End If
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub

Sub �I��͈͂ŏ�̃Z���l�ƈقȂ�ꍇ�̓Z���F�����F�ɂ���()
    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Dim i As Long, j As Long

    ' �I������Ă���͈͂��擾
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)

    For i = Left To Right
        For j = Top To bottom
            ' �}�[�W����Ă���Z���͖���
            If Not Cells(j, i).MergeCells Then
                ' �l���قȂ�����F��t����
                If j >= 2 And Len(Cells(j, i).Text) > 0 And Cells(j - 1, i).Text <> Cells(j, i).Text Then
                    Cells(j, i).Interior.color = 10092543  ' ���F
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub

Sub �I��͈͂ŏ�̃Z���Ɠ����Ȃ�t�H���g�F���O���[�ɂ���()
    Dim selectedRange As Range
    Dim cell As Range
    Dim previousCell As Range

    ' �I��͈͂��擾
    Set selectedRange = Selection

    ' 2�s�ڈȍ~�̊e�Z���ɑ΂��ď��������s
    For Each cell In selectedRange
        If cell.Row > 1 Then  ' 1�s�ڂ͖�������
            Set previousCell = cell.Offset(-1)   ' ��̃Z�����擾

            ' �Z�����u�����N�łȂ��ꍇ�ɏ��������s
            If Not IsEmpty(cell.Value) Then
                ' �Z���̒l����̃Z���Ɠ�������r
                If cell.Value = previousCell.Value Then
                    ' �Z���������ŕ\����Ă���ꍇ�͌v�Z���ʂ��r
                    If cell.HasFormula Then
                        If cell.Value = previousCell.Value Then
                            cell.Font.color = RGB(192, 192, 192)   ' �O���[�ɂ���
                        End If
                    Else
                        cell.Font.color = RGB(192, 192, 192)    ' �O���[�ɂ���
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Sub �I��͈͂ŏ�̃Z���Ɠ����Ȃ�t�H���g�F�𔖂�����()
    Dim selectedRange As Range
    Dim cell As Range
    Dim previousCell As Range
    Dim currentColor As Long
    Dim newColor As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    ' �I��͈͂��擾
    Set selectedRange = Selection

    ' 2�s�ڈȍ~�̊e�Z���ɑ΂��ď��������s
    For Each cell In selectedRange
        If cell.Row > 1 Then  ' 1�s�ڂ͖�������
            Set previousCell = cell.Offset(-1)   ' ��̃Z�����擾

            ' �Z�����u�����N�łȂ��ꍇ�ɏ��������s
            If Not IsEmpty(cell.Value) Then
                ' �Z���̒l����̃Z���Ɠ�������r
                If cell.Value = previousCell.Value Then
                    ' �Z���������ŕ\����Ă���ꍇ�͌v�Z���ʂ��r
                    If cell.HasFormula Then
                        If cell.Value = previousCell.Value Then
                            currentColor = cell.Font.color ' ���݂̐F���擾
                            redValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - (currentColor And 255)) * 3 / 4 + (currentColor And 255), 0), 255) ' �Ԃ̒l���v�Z
                            greenValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256) And 255)) * 3 / 4 + ((currentColor \ 256) And 255), 0), 255)  ' �΂̒l���v�Z
                            blueValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256 \ 256) And 255)) * 3 / 4 + ((currentColor \ 256 \ 256) And 255), 0), 255) ' �̒l���v�Z
                            newColor = RGB(redValue, greenValue, blueValue) ' �����F���v�Z
                            cell.Font.color = newColor   ' �F��ݒ�
                        End If
                    Else
                        currentColor = cell.Font.color ' ���݂̐F���擾
                        redValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - (currentColor And 255)) * 3 / 4 + (currentColor And 255), 0), 255) ' �Ԃ̒l���v�Z
                        greenValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256) And 255)) * 3 / 4 + ((currentColor \ 256) And 255), 0), 255)  ' �΂̒l���v�Z
                        blueValue = Application.WorksheetFunction.Min(WorksheetFunction.RoundUp((256 - ((currentColor \ 256 \ 256) And 255)) * 3 / 4 + ((currentColor \ 256 \ 256) And 255), 0), 255) ' �̒l���v�Z
                        newColor = RGB(redValue, greenValue, blueValue) ' �����F���v�Z
                        cell.Font.color = newColor   ' �F��ݒ�
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Sub �I��͈͂Ő��}�[�N�̃Z����������͐܂肽���ݐݒ�()
    Dim myRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    For Each myRange In Selection
        If myRange.Value = "��" Or myRange.Value = "��" Then
            If Not map.Exists(myRange.Column) Then
                Call map.Add(myRange.Column, True)
            End If
        End If
    Next

    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub �I��͈͂Ő��}�[�N�̃Z���̉��͐ԕ���()
    Dim myRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    For Each myRange In Selection
        If myRange.Value = "��" Or myRange.Value = "��" Then
            Cells(myRange.Row + 1, myRange.Column).Font.color = 255
        End If
    Next
End Sub

Sub �I��͈͂Œl�������Ă��Ȃ���͐܂肽���ݐݒ�()
    Dim myRange As Range
    Dim targetRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)

    If Top = bottom Then
        ' 1�s�����I������Ă��Ȃ��ꍇ�͑f���ɂ��͈̔͂�����������
        Set targetRange = Selection
    Else
        ' �����s�I������Ă���ꍇ��1�s�ڂ͖�������B���̃}�N���͕\�`���Ŏg���邱�Ƃ�z�肵�Ă��āA�w�b�_�s�͖����������̂ŁB
        Set targetRange = Range(Cells(Top + 1, Left), Cells(bottom, Right))
    End If

    ' �l�������Ă������L��
    For Each myRange In targetRange
        If Not IsEmpty(myRange.Value) Then
            If Not map.Exists(myRange.Column) Then
                Call map.Add(myRange.Column, True)
            End If
        End If
    Next

    ' �l�������Ă��Ȃ���͐܂肽����
    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub �I��͈͂Œl�������Ă��Ȃ���͐܂肽���ݐݒ�_�h��Ԃ��̂���Z���͏��O()
    Dim myRange As Range
    Dim targetRange As Range
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim Top As Long, bottom As Long, Left As Long, Right As Long
    Call func�I��͈͂̍��W���擾(Top, bottom, Left, Right)

    If Top = bottom Then
        ' 1�s�����I������Ă��Ȃ��ꍇ�͑f���ɂ��͈̔͂�����������
        Set targetRange = Selection
    Else
        ' �����s�I������Ă���ꍇ��1�s�ڂ͖�������B���̃}�N���͕\�`���Ŏg���邱�Ƃ�z�肵�Ă��āA�w�b�_�s�͖����������̂ŁB
        Set targetRange = Range(Cells(Top + 1, Left), Cells(bottom, Right))
    End If

    ' �l�������Ă������L��
    For Each myRange In targetRange
        If Not IsEmpty(myRange.Value2) Then
            ' �h��Ԃ�����Ă��Ȃ������ꍇ�̂�
            If myRange.Interior.colorIndex = xlNone Then
                If Not map.Exists(myRange.Column) Then
                    Call map.Add(myRange.Column, True)
                End If
            End If
        End If
    Next

    ' �l�������Ă��Ȃ���͐܂肽����
    For Each myRange In Selection
        If myRange.Columns.OutlineLevel = 1 Then
            If Not map.Exists(myRange.Column) Then
                myRange.Columns.Group
            End If
        End If
    Next
End Sub

Sub �I�𒆂̃Z���ʒu���N���b�v�{�[�h�ɃR�s�[()
    Dim address As String
    Dim sheetName As String
    Dim msg As String
    address = Selection.address
    address = Replace(address, "$", "")
    sheetName = ActiveSheet.Name

    msg = ActiveWorkbook.Name & vbLf & "'" & sheetName & "'!" & address

    PutClipBoard (msg)
    Application.StatusBar = msg  ' �f�o�b�O�p�Ɉꎞ�I�ɏ������B
End Sub

' �I��͈͂̃Z���̏������X�V����}�N��
' �I��͈͂̒��ň�ԍ���̃Z���̈ȉ��̏����𑼂̑I���Z���ɔ��f����
'   �E�Z���̐F(�w�i�F)
'   �E�����F
'   �E�r��
'   �E������
Sub ����Z���̏����𑼃Z���ɃR�s�[()
    Dim sourceCell As Range
    Dim targetRange As Range

    ' �I��͈͂̒��ň�ԍ���̃Z�����擾
    Set sourceCell = Selection.Cells(1)

    ' �I��͈͂��擾
    Set targetRange = Selection

    ' �������R�s�[����
    sourceCell.Copy

    ' �����𑼂̑I���Z���ɔ��f����
    targetRange.PasteSpecial xlPasteFormats

    ' �X�e�[�^�X�o�[�Ƀ��b�Z�[�W��\�����A5�b��ɏ���
    Application.StatusBar = "�������R�s�[���܂���"
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Sub �I��͈͂̉�������������()
    ' ��̕��̎������� (�I��͈͂Œ���)
    Selection.Columns.AutoFit

    ' �I��͈̗͂�����[�v���āA�������傫��������̂͌Œ蒷�ɒ���
    Const MAX_WIDTH As Double = 50 ' �ő�̗񕝂��`
    Dim col As Range
    For Each col In Selection.Columns
        If col.ColumnWidth > MAX_WIDTH Then
            col.ColumnWidth = MAX_WIDTH
        End If
    Next col

    Application.ScreenUpdating = True
End Sub

Attribute VB_Name = "BatchModule"
Option Explicit

Sub �V�[�g��̃I�u�W�F�N�g��S�폜()
    Dim rc As Long
    rc = MsgBox("Are you sure to delete all shapes?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If
    
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub

Sub �J�����g�V�[�g�̃R�����g��S�폜()
    Dim rc As Long
    rc = MsgBox("Are you sure to delete all comments?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If
    
    On Error Resume Next
    Cells.SpecialCells(xlCellTypeComments).ClearComments
End Sub

Sub �A�N�e�B�u�u�b�N�̃I�[�g�t�B���^��S�ĉ�������()
    Dim sh As Worksheet
    
    For Each sh In Worksheets
        If sh.AutoFilterMode Then
            If sh.AutoFilter.FilterMode Then
                sh.ShowAllData
            End If
        End If
    Next sh
End Sub

Sub B2�Z���̓��e�ŃV�[�g����ύX()
    Dim str As String: str = ActiveSheet.Cells(2, 2).Value
    Dim length As Long: length = Len(str)
    
    If length > O And length <= 31 Then
        ActiveSheet.Name = str
    End If
End Sub


Sub �J�����gBook�̖��O�̒�`���폜()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' �G���[�𖳎�
        n.Delete
    Next
End Sub

Sub �J�����gBook�̖��O�̒�`���폜_����ݒ�ȊO()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If InStr(n.Name, "Print_") = 0 Then
            On Error Resume Next ' �G���[�𖳎�
            n.Delete
        End If
    Next
End Sub

Sub �J�����gBook�̖��O�̒�`���폜_�G���[�̂�()
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' �G���[�𖳎�
               If InStr(n.Value, "=#") = 1 Then
            n.Delete
        End If
    Next
End Sub

Sub �J�����gBook�̖��O�̒�`�ŕʃu�b�N�̎Q�Ƃ����Ă�����̂��폜()
    Dim n As Name
    Dim count As Long: count = 0

    Application.StatusBar = False

    For Each n In ActiveWorkbook.Names
        On Error Resume Next ' �G���[�𖳎�

        If left(n.Value, 4) = "='\\" Or left(n.Value, 5) = "='C:\" Then
            count = count + 1
            n.Delete
        End If
    Next

    ' �X�e�[�^�X�p�̍X�V
    If count <> 0 Then
        Application.StatusBar "�ʃu�b�N�Q�Ƃ̖��O�F" & count & "�����폜���܂����B"
    Else
        Application.StatusBar = False
    End If
End Sub

Sub DIFF�K���̃V�[�g�𐮌`()
    Dim activeSheetBak As Worksheet: Set activeSheetBak = ActiveSheet
    Dim sh As Worksheet

    Application.ScreenUpdating = False

    For Each sh In Worksheets
        ' A1�Z����Diff�c�[���Ŏg���Ă���F��������Diff�V�[�g�Ɣ��f
        If sh.Range("A1").Interior.COLOR = 16711680 Then
            Call func_DIFF�`���̃V�[�g�𐮌`_1�V�[�g(sh)
        End If
    Next sh

    activeSheetBak.Activate
    Application.ScreenUpdating = True
End Sub

Sub �J�����gBook�̖��O�̒�`�ŕʃu�b�N�Q�Ƃ�����ӏ������o()
    Dim n As Name
    Dim jumpedFlg As Boolean: jumpedFlg = False
    Dim count As Long: count = 0
    Dim firstHit As String

    Application.StatusBar = False

    For Each n In ActiveWorkbook.Names
        On Error Resume Next '�G���[�𖳎�

        If left(n.Value, 4) = "='\\" Or left(n.Value, 5) = "='C:\" Then
            count = count + 1
            Debug.Print "--------------------------"
            Debug.Print n.Name
            Debug.Print n.Value
            Debug.Print n.Parent.CodeName
            Debug.Print n.Parent.Authore

            ' �ŏ��Ƀq�b�g�������O�ɃW�����v����
            If jumpedFlg = False Then
                Application.GoTo Reference:=n.Name
                jumpedFlg = True
                firstHit = "name:[" & n.Name & "] CodeName:[" & n.Parent.CodeName & "]"
            End If
        End If
    Next

    ' �X�e�[�^�X�o�[�̍X�V
    If jumpedFlg Then
        Application.StatusBar = "�ʃu�b�N�Q�Ƃ̖��O��" & count & "��������܂����B(first Hit ->" & firstHit & ")"
    Else
        Application.StatusBar = False
    End If
End Sub

Private Function func_DIFF�`���̃V�[�g�𐮌`_1�V�[�g(ByVal sh As Worksheet)
    sh.Select
    ' �E�B���h�E�g�̌Œ�
    sh.Rows("2:2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' �\���T�C�Y��75%��
    ActiveWindow.Zoom = 75

    ' �s�ԍ��������Ă���A���C��̃T�C�Y�𒲐�
    sh.Range("A:A,C:C").ColumnWidth = 5

    ' �t�@�C���̒��g�������Ă���B���D��̃T�C�Y�𒲐�
    sh.Range("B:B,D:D").ColumnWidth = 95

    ' DIFF��̐ݒ�A����
    sh.Range("E1").FormulaR1C1 = "DIFF"
    sh.Columns("E:E").ColumnWidth = 4

    ' ���ɃI�[�g�t�B���^���ݒ肳��Ă���ꍇ�͉���
    If sh.AutoFilterMode = True Then
        If sh.AutoFilter.FilterMode = True Then
            sh.ShowAllData
        End If
        sh.Rows("1:1").AutoFilter
    End If

    ' �I�[�g2�t�B���^�̐ݒ�
    sh.Rows("1:1").AutoFilter

    ' �����̂���s(���F�̍s)���t�B���^
    sh.Range(Range("A1").Cells(Rows.count, 3).End(xlUp)).AutoFilter Field:=4, Criteria1:=RGB(239, 203, 5), Operator:=xlFilterCellColor
    
    ' E1�Z����I��
    sh.Range("E1").Select
End Function



Sub �V�[�g��̃I�u�W�F�N�g��S�폜()
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
End Sub



Sub ���[�N�V�[�g�𖼑O�Ń\�[�g()
    Application.ScreenUpdating = False
    
    Dim i As Long, j As Long, cnt As Long
    Dim buf() As String, swap As String
    Dim selectedSheet As Worksheet
    
    ' ���X�J���Ă����V�[�g���L��
    Set selectedSheet = Application.ActiveWorkbook.ActiveSheet
    
    cnt = Application.ActiveWorkbook.Worksheets.count
    
    If cnt > 1 Then
        ReDim buf(cnt)
        
        ' ���[�N�V�[�g����z��ɓ����
        For i = 1 To cnt
            buf(i) = Application.ActiveWorkbook.Worksheets(i).Name
        Next i
        
        ' �z��̗v�f���\�[�g����
        For i = 1 To cnt
            For j = cnt To i Step -1
                If buf(i) > buf(j) Then
                    swap = buf(i)
                    buf(i) = buf(j)
                    buf(j) = swap
                End If
            Next j
        Next i
    End If
    
    ' ���[�N�V�[�g�̈ʒu����ёւ���
    Application.ActiveWorkbook.Worksheets(buf(1)).Move Before:=Application.ActiveWorkbook.Worksheets(1)
    
    For i = 2 To cnt
        Application.ActiveWorkbook.Worksheets(buf(i)).Move After:=Application.ActiveWorkbook.Worksheets(i - 1)
    Next i
    
    selectedSheet.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub �V�[�g���ꗗ���쐬����()
    Application.ScreenUpdating = False
    
    Const addSheetName As String = "�V�[�g���ꗗ(��������)"
    
    Dim i As Long
    Dim ws As Worksheet
    Dim flag As Boolean
    
    For Each ws In Application.ActiveWorkbook.Worksheets
        ' addSheetName�̖��̂̃V�[�g������������폜����
        If ws.Name = addSheetName Then
            Application.DisplayAlerts = False   ' �폜���̌x�����b�Z�[�W�͔�\��
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next ws
        
    Application.ActiveWorkbook.Sheets.Add Before:=Application.ActiveWorkbook.Worksheets(1)
    Application.ActiveWorkbook.ActiveSheet.Name = addSheetName
    
    For i = 2 To Application.ActiveWorkbook.Sheets.count
        Application.ActiveWorkbook.ActiveSheet.Cells(i - 1, "A").Value = Application.ActiveWorkbook.Sheets(i).Name
    Next i

    Application.ScreenUpdating = True
End Sub

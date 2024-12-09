Attribute VB_Name = "BatchModule"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub �V�[�g��̃I�u�W�F�N�g��S�폜()
    Dim rc As Integer
    rc = MsgBox("Are you sure to delete all shapes on current sheet?", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
        ActiveSheet.Shapes.SelectAll
        Selection.Delete
    End If
End Sub

Sub �V�[�g��̃I�u�W�F�N�g�̏����ύX_�Z���ɂ��킹�Ĉړ��⃊�T�C�Y������()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMoveAndSize ' �Z���폜��ړ��ɍ��킹�Ĉړ����A���T�C�Y���s��
End Sub

Sub �V�[�g��̃I�u�W�F�N�g�̏����ύX_�Z���ɂ��킹�Ĉړ��⃊�T�C�Y���Ȃ�()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlFreeFloating ' �Z���폜��ړ��ɍ��킹�����T�C�Y�A�ړ����s��Ȃ�
End Sub

Sub �V�[�g��̃I�u�W�F�N�g�̏����ύX_�Z���ɂ��킹�Ĉړ����邪���T�C�Y���Ȃ�()
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMove ' �Z���폜��ړ��ɍ��킹�Ĉړ�����
End Sub

Sub �J�����g�V�[�g�̃R�����g��S�폜()
    Dim rc As Integer
    rc = MsgBox("Are you sure to delete all comments on current sheet?", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
        On Error Resume Next
        Cells.SpecialCells(xlCellTypeComments).ClearComments
    End If
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
    Dim newSheetName As String
    Dim temp As String
    Dim i As Integer
    
    If length > 0 And length <= 31 Then
        newSheetName = str
    ElseIf length > 31 Then
        newSheetName = left(str, 31)
    End If

    If ActiveSheet.Name <> newSheetName And SheetExists(newSheetName) Then
        If Len(newSheetName) > 28 Then
            newSheetName = left(newSheetName, 28)
        End If
        
        For i = 2 To 10
            temp = newSheetName & "(" & i & ")"
            If ActiveSheet.Name = temp Or Not SheetExists(temp) Then
                newSheetName = temp
                Exit For
            End If
        Next i
    End If
    
    If Len(newSheetName) > 0 Then
        ActiveSheet.Name = newSheetName
    End If
End Sub

Function SheetExists(shtName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(shtName) Is Nothing
    On Error GoTo 0
End Function


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
        Application.StatusBar = "�ʃu�b�N�Q�Ƃ̖��O�F" & count & "�����폜���܂����B"
    Else
        Application.StatusBar = False
    End If
End Sub

Sub DIFF�����̃V�[�g�𐮌`()
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

Sub DIFF�����̃V�[�g�𐮌`_�A�N�e�B�u�V�[�g�̂�()
    Application.ScreenUpdating = False
    Call func_DIFF�`���̃V�[�g�𐮌`_1�V�[�g(ActiveSheet)
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
                Application.Goto Reference:=n.Name
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
    Dim YELLOW As Long: YELLOW = RGB(239, 203, 5)
    Dim GRAY As Long: GRAY = RGB(192, 192, 192)
    Dim LIGHT_PINK As Long: LIGHT_PINK = RGB(240, 192, 192)
    Dim PINK As Long: PINK = RGB(239, 119, 116)
    
    Dim i As Long
    Dim sheetName As String
    Dim bottom As Long
    Dim tmp As Long

    sh.Select
    ' �E�B���h�E�g�̌Œ�
    sh.Rows("2:2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' �\���T�C�Y��ύX
    ActiveWindow.Zoom = 85

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
    bottom = Cells(Rows.count, 1).End(xlUp).Row   ' A��̍ŏI�s
    tmp = Cells(Rows.count, 3).End(xlUp).Row ' C��̍ŏI�s
    If bottom < tmp Then
        bottom = tmp
    End If
    For i = 2 To bottom
        If Cells(i, 2).Interior.COLOR = GRAY Or Cells(i, 4).Interior.COLOR = GRAY Or _
           Cells(i, 2).Interior.COLOR = YELLOW Or Cells(i, 4).Interior.COLOR = YELLOW Or _
           Cells(i, 2).Interior.COLOR = LIGHT_PINK Or Cells(i, 4).Interior.COLOR = LIGHT_PINK Or _
           Cells(i, 2).Interior.COLOR = PINK Or Cells(i, 4).Interior.COLOR = PINK Then
            ' ���O�̍s����������Ɣ��肳��Ă�����A����ElseIf�̔���������ɂ��̍s����������Ɣ���B
            ' �A�����Ĉ��ɂȂ��Ă���Ȃ�A���̍s�������������Ɣ��肵�Ă��g�����肪�����̂ŁB
            If Cells(i - 1, 5).Value = "��" Then
                Cells(i, 5).Value = "��"
            ' ��s�͏��O�B���ƁAimport�Ŏn�܂�s�����O�B
            ElseIf Not (Len(Cells(i, 2).Value) = 0 And Len(Cells(i, 4).Value) = 0) And _
              Not (CommonModule.regularExpressionTest(Cells(i, 2).Value, "^import .+;", False)) And _
              Not (CommonModule.regularExpressionTest(Cells(i, 4).Value, "^import .+;", False)) Then
                Cells(i, 5).Value = "��"
           End If
        End If
    Next i
    
    ' E1�Z����I��
    sh.Range("E1").Select
End Function

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

Sub �V�[�g���ꗗ���N���b�v�{�[�h�ɃR�s�[()
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetNames As String: sheetNames = ""

    For Each ws In Application.ActiveWorkbook.Worksheets
        sheetNames = sheetNames & ws.Name & vbCrLf
    Next ws
    
    Call PutClipBoard(sheetNames)
End Sub

Sub �S�V�[�g�{��100�p�[�Z���g�ɂ��Đ擪�Z���I��()
    Dim sht   As Worksheet              ' �������̃��[�N�V�[�g
    Dim shtVisible                      ' �\���\�ȃ��[�N�V�[�g
    Dim iRow, iCol                      ' �c�A�����W
    Dim oFilterStatus As AutoFilter     ' �I�[�g�t�B���^���
    Dim oRangeFilter As Range           ' �I�[�g�t�B���^�ݒ�
    Dim zoomRc As Integer
    Dim zoomMsgBoxConducted As Boolean: zoomMsgBoxConducted = False

    Application.ScreenUpdating = True

    For Each sht In Sheets
        If (IsEmpty(shtVisible) = True) And (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            Set shtVisible = sht
        End If

        ' �V�[�g���\������Ă���ꍇ
        If (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            sht.Select

            ' 85%�ȉ��̂Ƃ���85%�ɂ���
            If ActiveWindow.Zoom <= 85 Then
                ActiveWindow.Zoom = 85
            Else
                ActiveWindow.Zoom = 100
            End If

            ' �E�C���h�E�g�̌Œ肪����Ă���ꍇ
            If ActiveWindow.FreezePanes = True Then
                iRow = ActiveWindow.SplitRow + 1
                iCol = ActiveWindow.SplitColumn + 1
                Cells(iRow + 1, iCol + 1).Activate
            End If

            Set oFilterStatus = sht.AutoFilter
            ' �I�[�g�t�B���^���ݒ肳��Ă���ꍇ
            If Not oFilterStatus Is Nothing Then
                ' �t�B���^���|�����Ă���ꍇ
                If oFilterStatus.FilterMode = True Then
                    ' �t�B���^���|�����Ă���s�̐擪��I��
                    Set oRangeFilter = Range("A1").CurrentRegion
                    Set oRangeFilter = Application.Intersect(oRangeFilter, oRangeFilter.Offset(1, 0))
                    Set oRangeFilter = oRangeFilter.SpecialCells(xlCellTypeVisible)
                    Range("A" & CStr(oRangeFilter.Row)).Select
                End If
            End If
            
            sht.Range("A1").Select
        End If
    Next
    
    shtVisible.Select

End Sub

Sub �S�V�[�g���y�[�W�v���r���[_�g����()
    Dim sht As Worksheet ' �������̃��[�N�V�[�g
    Dim shtVisible      ' �\���\�ȃ��[�N�V�[�g
    Dim iRow, iCol ' �c�A�����W
    Dim oFilterStatus As AutoFilter  ' �I�[�g�t�B���^���
    Dim oRangeFilter As Range ' �I�[�g�t�B���^�ݒ�
    Dim zoomRc As Integer
    Dim zoomMsgBoxConducted As Boolean: zoomMsgBoxConducted = False
    Dim rc As Integer
    
    rc = MsgBox("Are you sure to change all sheets format ?", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = True
    
    For Each sht In Sheets
        If (IsEmpty(shtVisible) = True) And (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            Set shtVisible = sht
        End If

        ' �V�[�g���\������Ă���ꍇ
        If (sht.Visible <> xlSheetHidden) And (sht.Visible <> xlSheetVeryHidden) Then
            sht.Select
            ActiveWindow.View = xlPageBreakPreview ' ���y�[�W�v���r���[
            
            ' 85%�ȉ��̂Ƃ���85%�ɂ���
            If ActiveWindow.Zoom <= 85 Then
                ActiveWindow.Zoom = 85
            Else
                ActiveWindow.Zoom = 100
            End If
            ActiveWindow.DisplayGridlines = False ' �g������

            ' �E�C���h�E�g�̌Œ肪����Ă���ꍇ
            If ActiveWindow.FreezePanes = True Then
                iRow = ActiveWindow.SplitRow + 1
                iCol = ActiveWindow.SplitColumn + 1
                Cells(iRow + 1, iCol + 1).Activate
            End If

            Set oFilterStatus = sht.AutoFilter
            ' �I�[�g�t�B���^���ݒ肳��Ă���ꍇ
            If Not oFilterStatus Is Nothing Then
                ' �t�B���^���|�����Ă���ꍇ
                If oFilterStatus.FilterMode = True Then
                    ' �t�B���^���|�����Ă���s�̐擪��I��
                    Set oRangeFilter = Range("A1").CurrentRegion
                    Set oRangeFilter = Application.Intersect(oRangeFilter, oRangeFilter.Offset(1, 0))
                    Set oRangeFilter = oRangeFilter.SpecialCells(xlCellTypeVisible)
                    Range("A" & CStr(oRangeFilter.Row)).Select
                End If
            End If

            sht.Range("A1").Select
        Else
            sht.Visible = xlSheetVisible ' �V�[�g��\��
            sht.Select
            
            ActiveWindow.DisplayGridlines = False ' �g������
            
            sht.Visible = xlSheetHidden ' �V�[�g���\��
        End If
    Next

    shtVisible.Select

End Sub

Sub �S�V�[�g�w��{���ɕύX()
    Dim shOriginalSelected As Worksheet
    Dim sh As Worksheet
    Dim strScale As String
    Dim nScale As Long
    
    strScale = InputBox("�{�����w�肵�Ă�������(""**%""�̐������������)")
    If strScale = "" Then Exit Sub
    nScale = CLng(strScale)

    Application.ScreenUpdating = False
    
    Set shOriginalSelected = ActiveSheet

    For Each sh In Sheets
        If (sh.Visible <> xlSheetHidden) And (sh.Visible <> xlSheetVeryHidden) Then
            sh.Select
            ActiveWindow.Zoom = nScale
        End If
    Next
    
    shOriginalSelected.Activate

    Application.ScreenUpdating = True
End Sub

Sub �N���b�v�{�[�h�̉��()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
    'Application.CutCopyMode = False
End Sub

Sub �J�����g�V�[�g�̑S�Z����MeiryoUI��()
    Cells.Font.Name = "Meiryo UI"
End Sub

Sub �V�[�g�\���ݒ�_�F���t���ĂȂ��V�[�g�͔�\����()
    Dim ws As Worksheet
    Dim containColorTab As Boolean
    Application.ScreenUpdating = False
    
    containColorTab = False
    For Each ws In Worksheets
        If ws.Tab.ColorIndex <> xlNone Then
            containColorTab = True
            Exit For
        End If
    Next ws
    
    If containColorTab Then
        For Each ws In Worksheets
            If ws.Tab.ColorIndex = xlNone Then
                If ws.Visible = xlSheetVisible Then
                    ws.Visible = False
                End If
            End If
        Next ws
        Application.StatusBar = False
    Else
        Application.StatusBar = "�F�t���V�[�g�������̂ŏ������܂���ł���"
    End If

    Application.ScreenUpdating = True
End Sub

Sub �V�[�g�\���ݒ�_�S�V�[�g�\��()
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each ws In Worksheets
        ws.Visible = True
    Next ws
    Application.ScreenUpdating = True

    Application.StatusBar = False
End Sub

Sub �A�N�e�B�u�u�b�N�̈��������S�ĉ���()
    Dim sh As Worksheet
    Dim rc As Integer
    Dim preCheckResult As String

    ' �܂��͉������Əc�����̂ǂ���ɂȂ��Ă��邩�`�F�b�N
    For Each sh In Worksheets
        ' ������
        If sh.PageSetup.Orientation = xlLandscape Then
            preCheckResult = preCheckResult & "Landscape"
        ' �c����
        Else
            preCheckResult = preCheckResult & "Portrait"
        End If
        
        preCheckResult = preCheckResult & " : " & sh.Name & vbCrLf
    Next sh
    
    rc = MsgBox(preCheckResult & vbCrLf & "Are you sure to set LANDSCAPE mode on all sheets?", vbYesNo + vbQuestion)

    If rc = vbYes Then
        For Each sh In Worksheets
            If sh.PageSetup.Orientation = xlPortrait Then
                sh.PageSetup.Orientation = xlLandscape
            End If
        Next sh
    End If
End Sub

Sub �A�N�e�B�u�u�b�N��CSV�`���ŕۑ�()
    Application.ScreenUpdating = False

    Dim cnt As Long: cnt = 1
    Dim bookName As String
    Dim directoryPath As String
    Dim activeBookFullPath As String

    bookName = left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    activeBookFullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name

    directoryPath = ActiveWorkbook.Path & "\" & bookName
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
    End If

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ActiveWorkbook.SaveAs Filename:=directoryPath & "\" & cnt & "_" & ws.Name & ".csv", FileFormat:=xlCSV
        cnt = cnt + 1
    Next ws
    
    ' �A�N�e�B�u�u�b�N��CSV�`���̖��O�ɂȂ��Ă���̂ŁA��x�t�@�C������čēx�J���B
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open activeBookFullPath

    Application.ScreenUpdating = True
    MsgBox "�A�N�e�B�u�u�b�N��CSV�`���ŕۑ����܂���"
End Sub

Sub �A�N�e�B�u�V�[�g��CSV�`���ŕۑ�()
    Application.ScreenUpdating = False
    
    Dim bookName As String
    Dim directoryPath As String
    Dim activeBookFullPath As String
    
    bookName = left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
    activeBookFullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name

    directoryPath = ActiveWorkbook.Path & "\" & bookName
    If Dir(directoryPath, vbDirectory) = "" Then
        MkDir directoryPath
    End If

    ActiveWorkbook.SaveAs Filename:=directoryPath & "\" & ActiveSheet.Name & ".csv", FileFormat:=xlCSV
    
    ' �A�N�e�B�u�u�b�N��CSV�`���̖��O�ɂȂ��Ă���̂ŁA��x�t�@�C������čēx�J���B
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open activeBookFullPath
    
    Application.ScreenUpdating = True
    MsgBox "�A�N�e�B�u�V�[�g��CSV�`���ŕۑ����܂���"
End Sub

Sub �V�[�g��̉摜�ɘg����ǉ�����()
    Const MAX_IMAGE_COUNT As Long = 100

    Dim imageCount As Long
    Dim continueProcessing As Boolean

    ' �摜�̐����J�E���g
    imageCount = ActiveSheet.Shapes.count

    ' �摜�̐�����萔�𒴂��Ă���ꍇ�A�x����\�����ď������s�̊m�F���擾
    If imageCount > MAX_IMAGE_COUNT Then
        continueProcessing = MsgBox("�摜�̐��� " & imageCount & " ����܂��B�������Ԃ������Ȃ�\��������܂����A���s���܂���?", vbQuestion + vbYesNo) = vbYes
    Else
        continueProcessing = True
    End If

    ' �����𑱍s����ꍇ�A�摜�ɘg����ǉ�
    If continueProcessing Then
        Dim shape As shape
        
        For Each shape In ActiveSheet.Shapes
            ' �摜�݂̂ɘg����ǉ�
            If shape.Type = msoPicture Then
                shape.Line.Weight = 0.5   ' �g���̑�����ݒ�
                shape.Line.ForeColor.RGB = RGB(0, 0, 0)  ' �g���̐F�����ɐݒ�
            End If
        Next shape

        Application.StatusBar = "�摜�ɘg����ǉ����܂����B"

        ' �X�e�[�^�X�o�[�̕\����5�b��ɏ���
        Application.OnTime Now + TimeValue("00:00:05"), "CommonModule.ResetStatusBar"
    Else
        MsgBox "�����𒆎~���܂����B", vbInformation
    End If
End Sub

Sub �u�b�N��̑S�V�[�g�̉摜�ɘg����ǉ�����()
    Const MAX_IMAGE_COUNT As Long = 100

    Dim imageCount As Long
    Dim continueProcessing As Boolean
    
    Dim shape As shape
    Dim sh As Worksheet

    For Each sh In Worksheets
        For Each shape In sh.Shapes
            ' �摜�݂̂ɘg����ǉ�
            If shape.Type = msoPicture Then
                shape.Line.Weight = 0.5   ' �g���̑�����ݒ�
                shape.Line.ForeColor.RGB = RGB(0, 0, 0) ' �g���̐F�����ɐݒ�
            End If
        Next shape
    Next sh

    Application.StatusBar = "�摜�ɘg����ǉ����܂����B"

    ' �X�e�[�^�X�o�[�̕\����5�b��ɏ���
    Application.OnTime Now + TimeValue("00:00:05"), "CommonModule.ResetStatusBar"
End Sub


Sub �I�𒆂̃V�[�g�̗񕝂𑵂���()
    ' �I�𒆂̃V�[�g���擾
    Dim selectedSheet As Worksheet
    Set selectedSheet = ActiveSheet

    Dim lastColumn As Long
    lastColumn = selectedSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Dim sh As Worksheet
    Dim i As Long
    For Each sh In ActiveWindow.SelectedSheets
        If Not selectedSheet Is sh Then
            ' �e�V�[�g�̗񕝂̑���
            For i = 1 To lastColumn
                sh.Columns(i).ColumnWidth = selectedSheet.Columns(i).ColumnWidth
            Next i
        End If
    Next sh

    ' �X�e�[�^�X�o�[�ւ̏o�͂ƃN���A
    Application.StatusBar = "�������������܂����B"
    Application.DisplayStatusBar = True

    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Sub �I�𒆂̃V�[�g�̗񕝂ƃt�H���g�𑵂���()
    ' �I�𒆂̃V�[�g���擾
    Dim selectedSheet As Worksheet
    Set selectedSheet = ActiveSheet

    Dim lastColumn As Long
    lastColumn = selectedSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    ' �t�H���g�̑���
    Dim activeFontName As String
    On Error Resume Next
    activeFontName = selectedSheet.Cells(1, 1).Font.Name
    On Error GoTo 0

    Dim sh As Worksheet
    Dim i As Long
    For Each sh In ActiveWindow.SelectedSheets
        If Not selectedSheet Is sh Then
            ' �e�V�[�g�̗񕝂̑���
            For i = 1 To lastColumn
                sh.Columns(i).ColumnWidth = selectedSheet.Columns(i).ColumnWidth
            Next i
            
            If activeFontName <> "" Then
                sh.Cells.Font.Name = activeFontName
            End If
        End If
    Next sh

    ' �X�e�[�^�X�o�[�ւ̏o�͂ƃN���A
    Application.StatusBar = "�������������܂����B"
    Application.DisplayStatusBar = True

    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

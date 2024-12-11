Attribute VB_Name = "SearchModule"
Option Explicit

' --------------------------------------------------------------------------
' �ȉ��̎Q�Ɛݒ肪�K�v�ł��B
' �ݒ�́A[�c�[��]��[�Q�Ɛݒ�]�ŁB
' <���K�\������>
' "Microsoft VBScript Regular Expressions 5.5"
' --------------------------------------------------------------------------

Private Const COLOR_DEFAULT = -1

' --------------------------------------------------------------------------
' strPattern            ���������� (���K�\���Ŏw��)
' sh                    �Ώۂ̃��[�N�V�[�g
' ignoreCase            �啶���Ə���������ʂ���ꍇ��False�A��ʂ��Ȃ��ꍇ��True
' color_sheet           �����q�b�g�����Ƃ��ɃV�[�g�̐F��ύX����ꍇ�͐F���w��B�ύX���Ȃ��ꍇ��-1���w��B
' color_cell            �����q�b�g�����Ƃ��ɃZ���̐F�̓h��Ԃ���ύX����ꍇ�͐F���w��B�ύX���Ȃ��ꍇ��-1���w��B
' color_font            �����q�b�g�����Ƃ��ɊY���ӏ��̃t�H���g�̐F��ύX����ꍇ�͐F���w��B�ύX���Ȃ��ꍇ��-1���w��B
' regexSearch           ���K�\���Ō������邩�ǂ����BTrue�̏ꍇ�͐��K�\���Ō�������
' blnCellColorFlg       �����q�b�g�����ӏ��̃Z����h��Ԃ���
' blnFontColorFlg       �����q�b�g�����ӏ��̕����F��ύX���邩
' blnFontColorResetFlg  �����q�b�g�����ӏ��̕����F��ύX����ꍇ�A���Z�����̕��������Ń��Z�b�g���Ă���F�����邩(true: ���Z�b�g����, false: ���Z�b�g���Ȃ�)
' boldflag              �����q�b�g�����ӏ��𑾎��ɂ��邩���Ȃ����BTrue�̏ꍇ�����ɂ���
' underlineflag         �����q�b�g�����ӏ��ɉ����������������Ȃ����BTrue�̏ꍇ�͉���������
' strikethroughflag     �����q�b�g�����ӏ��Ɏ��������������������Ȃ����BTrue�̏ꍇ�͉���������
' markTopFlag           �����q�b�g�����ӏ��̗�̏�Ɂ���t���邩�t���Ȃ��� (true:�t����, false:�t���Ȃ�)
' jumpFirstHitCell      �����q�b�g�����Ƃ��ɤ�ŏ��Ƀq�b�g�����Z���ɃW�����v�����邩�Ƃ����True�̏ꍇ�̓W�����u����
' targetedSelectedCell  �����Ώۂ͈̔͂��A�I�������Z���Ɍ��肷��ꍇ�ATrue���Z�b�g����
' formatAndFormula      �����A�����𔽉f���Č������邩 (true:���f���������Ō���, false:���f���Ȃ������Ō���)�B�f�t�H���g��false�B
' --------------------------------------------------------------------------
Public Function func�����q�b�g���������������\��(ByVal strPattern As String, _
                        ByVal sh As Worksheet, _
                        ByVal ignoreCase As Boolean, _
                        ByVal color_sheet As Long, _
                        ByVal color_cell As Long, _
                        ByVal color_font As Long, _
                        Optional regexSearch As Boolean = True, _
                        Optional blnCellColorFlg As Boolean = True, _
                        Optional blnFontColorFlg As Boolean = True, _
                        Optional blnFontColorResetFlg As Boolean = False, _
                        Optional boldflag As Boolean = False, _
                        Optional underlineflag As Boolean = False, _
                        Optional strikethroughflag As Boolean = False, _
                        Optional markTopFlag As Boolean = False, _
                        Optional jumpFirstHitCell As Boolean = False, _
                        Optional targetedSelectedCell As Boolean = False, _
                        Optional formatAndFormula As Boolean = False _
                    ) As String

    Dim reg As New RegExp
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim startPos As Long
    Dim topRow As Long ' �擪�Ɂ��}�[�N������悤�ɁA�I���Z���̈�ԏ��Y���W���L������ϐ�
    Dim iLen
    Dim r As Range
    Dim iPosition
    Dim i
    Dim count As Long: count = 0
    Dim resultMessage As String: resultMessage = ""
    Dim targetRange As Range
    Dim cellStr As String
    
    ' ����������
    iLen = Len(strPattern)
    If iLen = 0 Then
        Exit Function
    End If
    
    If markTopFlag Then
        topRow = Rows.count
        For Each r In Selection
            If r.Row < topRow Then
                topRow = r.Row
                If r.Row < topRow Then
                    If topRow = 1 Then
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    
    If targetedSelectedCell Then
        Set targetRange = Selection
    Else
        Set targetRange = sh.UsedRange
    End If
    
    ' ���K�\���̏����ݒ�B
    reg.Global = True ' ������̍Ō�܂Ō���(True:����AFalse:���Ȃ�)
    reg.ignoreCase = ignoreCase ' �啶���Ə���������ʂ���ꍇ��False�A��ʂ��Ȃ��ꍇ��True
    reg.Pattern = strPattern

    ' �V�[�g�̐F���N���A�[
    If color_sheet <> COLOR_DEFAULT Then
        sh.Tab.colorIndex = xlNone
    End If

    count = 0
    
    ' �͈͂�1�Z�������[�v
    For Each r In targetRange
        If Not IsError(r.Value) Then
            If r.Value <> vbNullString Then
                ' ���K�\���Ō�������ꍇ
                If regexSearch Then
                    iPosition = 0
                    
                    ' �Z�������񂩂琳�K�\���ł̌������s��
                    If formatAndFormula Then
                        cellStr = r.Value
                    Else
                        cellStr = r.Text
                    End If
                    Set oMatches = reg.Execute(cellStr)
                    
                    ' �����Ō��������ӏ��̐������[�v
                    For i = 0 To oMatches.count - 1
                        ' ���������ꍇ�A�V�[�g�̐F��ύX
                        If color_sheet <> COLOR_DEFAULT Then
                            sh.Tab.color = color_sheet
                        End If
    
                        ' �������������J�E���g
                        count = count + 1
                        
                        ' ���������ӏ����擾
                        Set oMatch = oMatches.Item(i)
                        
                        ' ������v�̐擪�ʒu���擾
                        iPosition = oMatch.FirstIndex
    
                        ' ������v�����񒷂��擾
                        iLen = oMatch.length
                        
                        If i = 0 Then
                            If blnCellColorFlg Then
                                ' �Z���̓h��Ԃ�
                                r.Interior.color = color_cell
                            End If
                            
                            If blnFontColorResetFlg Then
                                ' �������܂ރZ���̏ꍇ�͕����F�̕ύX�͂��Ȃ�
                                If Not r.HasFormula Then
                                    ' �Z�����̕����F�������Ǎ��ɂ���
                                    r.Font.color = 0
                                End If
                            End If
                            
                            ' �����q�b�g�����Z���Ɉړ�
                            If jumpFirstHitCell Then
                                If count = 1 Then
                                    r.Activate
                                End If
                            End If
                        End If
                        
                        ' �������܂܂Ȃ��Z���̏ꍇ�͕����F�⑾�����̕ύX������
                        If Not r.HasFormula Then
                            ' ������v�����̃t�H���g��ύX
                            If boldflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Bold = True   ' ����
                            End If
                            If underlineflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Underline = True    ' ����������
                            End If
                            If strikethroughflag Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.Strikethrough = True    ' ��������������
                            End If
                            If blnFontColorFlg Then
                                r.Characters(Start:=iPosition + 1, length:=iLen).Font.color = color_font    ' �t�H���g�F
                            End If
                        End If
                        
                        ' �擪�Z���Ɂ��}�[�N��t����ꍇ
                        If (markTopFlag And (topRow > 1)) Then
                            If Cells(topRow - 1, r.Column).Value = "" Then
                                Cells(topRow - 1, r.Column).Value = "��"
                            End If
                        End If
                    Next
                ' �ʏ�̌���������ꍇ(���K�\���ł͂Ȃ��ꍇ)
                Else
                    startPos = 1    ' �������ڂ��猟�����邩
                    iPosition = -1  ' �������ڂŃq�b�g�������B�����l�͂Ƃ肠����-1�ŁB
                    i = 0           ' ���̃Z�����ł�������������
                    Do
                        If formatAndFormula Then
                            cellStr = r.Value
                        Else
                            cellStr = r.Text
                        End If
                    
                        ' �啶������������ʂ��Ȃ��ꍇ
                        If ignoreCase Then
                            ' �e�L�X�g���[�h�Ŕ�r����(�啶���E����������ʂ��Ȃ��A���p�E�S�p����ʂ��Ȃ�)
                            iPosition = InStr(startPos, cellStr, strPattern, vbTextCompare)
                        ' �啶������������ʂ���ꍇ
                        Else
                            ' �o�C�i�����[�h�Ŕ�r����(�啶���E����������ʂ���A���p�E�S�p����ʂ���)
                            iPosition = InStr(startPos, cellStr, strPattern, vbBinaryCompare)
                        End If
                        
                        ' ���������ꍇ
                        If iPosition > 0 Then
                            ' �������������J�E���g
                            count = count + 1
                            i = i + 1
                            
                            ' ���̃V�[�g���ŏ��߂ăq�b�g�����ꍇ
                            If count = 1 Then
                                ' ���������ꍇ�A�V�[�g�̐F��ύX
                                If color_sheet <> COLOR_DEFAULT Then
                                    sh.Tab.color = color_sheet
                                End If
                                
                                ' �����q�b�g�����Z���Ɉړ�
                                If jumpFirstHitCell Then
                                    r.Activate
                                End If
                            End If
                            
                            ' ���̃Z�����ŏ��߂ăq�b�g�����ꍇ
                            If i = 1 Then
                                If blnCellColorFlg Then
                                    ' �Z���̓h��Ԃ�
                                    r.Interior.color = color_cell
                                End If
                                
                                ' �������܂ރZ���̏ꍇ�͕����F�̕ύX�͂��Ȃ�
                                If Not r.HasFormula Then
                                    If blnFontColorResetFlg Then
                                        ' �Z�����̕����F�������Ǎ��ɂ���
                                        r.Font.color = 0
                                    End If
                                End If
                            End If
                            
                            ' �������܂܂Ȃ��Z���̏ꍇ�͕����F�⑾�����̕ύX������
                            If Not r.HasFormula Then
                                ' ������v�����̃t�H���g��ύX
                                If boldflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Bold = True   ' ����
                                End If
                                If underlineflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Underline = True    ' ����
                                End If
                                If strikethroughflag Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.Strikethrough = True ' ��������
                                End If
                                If blnFontColorFlg Then
                                    r.Characters(Start:=iPosition, length:=iLen).Font.color = color_font     ' �t�H���g�F
                                End If
                            End If
                            
                            ' �擪�Z���Ɂ��}�[�N��t����ꍇ
                            If (markTopFlag And (topRow > 1)) Then
                                If Cells(topRow - 1, r.Column).Value = "" Then
                                    Cells(topRow - 1, r.Column).Value = "��"
                                End If
                            End If
                            
                            startPos = iPosition + iLen
                        End If
                    Loop While iPosition <> 0
                End If
            End If
        End If
    Next
    
    If count <> 0 Then
        func�����q�b�g���������������\�� = sh.Name & ":" & count & "��, "
    End If

End Function

' TODO: �ȉ�Function�͖��g�p�H�Ȃ�폜���Ă����������B
' --------------------------------------------------------------------------
' a_sht                 ���[�N�V�[�g
' a_sPattern            �����p�^�[��
' a_bIgnoreCase         �啶���������̋�ʁiTrue�F��ʂ��Ȃ��AFalse�F��ʂ���j
' a_bFindReplace = True �����ƒu���̂ǂ��炩�iTrue�F�����AFalse�F�u���j
' a_sReplace = ""       �u��������
' --------------------------------------------------------------------------
Function funcFindCellRegExp(a_sht As Worksheet, a_sPattern As String, a_bIgnoreCase As Boolean, Optional a_bFindReplace As Boolean = True, Optional a_sReplace As String = "") As Range
    Dim reg         As New RegExp       '// ���K�\���N���X
    Dim iLen                            '// ������v������
    Dim r           As Range            '// �I���Z���͈͂̏������̂P�Z��
    Dim i                               '// ���[�v�J�E���^
    Dim bResult     As Boolean          '// ��������
    Dim rPre        As Range            '// �A�N�e�B�u�Z������̃Z���ň�v�����Z��
    Dim rFind       As Range            '// ������v�Z��
    
    '// ���������񂪖��ݒ�̏ꍇ
    iLen = Len(a_sPattern)
    If iLen = 0 Then
        Set funcFindCellRegExp = Nothing
        Exit Function
    End If
    
    '// ���K�\���̏����ݒ�
    reg.Global = True               '// ������̍Ō�܂Ō����iTrue�F����AFalse�F���Ȃ��j
    reg.ignoreCase = a_bIgnoreCase  '// �啶���������̋�ʁiTrue�F����AFalse�F���Ȃ��j
    reg.Pattern = a_sPattern        '// �������鐳�K�\���p�^�[��
    
    '// �Z���͈͂��P�Z�������[�v
    For Each r In a_sht.UsedRange
        '// �Z�������񂩂琳�K�\���ł̌������s��
        bResult = reg.Test(r.Value)
        
        '// �����Ɉ�v���Ȃ������ꍇ
        If bResult = False Then
            GoTo CONTINUE
        End If
        
        '// �ȉ������Ɉ�v�����ꍇ
        
        Debug.Print r.Address(False, False)
        
        '// ��Z���ł̌�����v�Ō��������Z�����܂������ꍇ
        If rPre Is Nothing Then
            '// ���݌������Ă���Z����ݒ�
            Set rPre = Range(r.Address)
        End If
        
        '// ���[�v�̃Z�����A�N�e�B�u�Z������ɂ���ꍇ
        If (r.Row < ActiveCell.Row) Then
            GoTo CONTINUE
        '// ���[�v�̃Z�����A�N�e�B�u�Z���Ɠ����s�����ǉE�ɂ���ꍇ
        ElseIf (r.Row = ActiveCell.Row) And (r.Column <= ActiveCell.Column) Then
            GoTo CONTINUE
        '// ���[�v�̃Z�����A�N�e�B�u�Z�����E���ɂ���ꍇ
        Else
            '// ������v�Z�������ݒ�̏ꍇ
            If rFind Is Nothing Then
                Set rFind = Range(r.Address)
            End If
        End If
        
CONTINUE:
    Next
    
    '// ���������ꍇ
    If Not rFind Is Nothing Then
        Set funcFindCellRegExp = rFind
        'rFind.Select
    '// �A�N�e�B�u�Z�����㑤�Ō��������ꍇ
    ElseIf Not rPre Is Nothing Then
        Set funcFindCellRegExp = rPre
        'rPre.Select
    '// ������Ȃ������ꍇ
    Else
        Set funcFindCellRegExp = Nothing
        'Call MsgBox("�����Ώۂ�������܂���", vbExclamation, "���K�\������")
        Exit Function
    End If
    
    '// �u���̏ꍇ
    If a_bFindReplace = False Then
        '// �A�N�e�B�u�Z���̕������u��
        ActiveCell.Value = reg.Replace(ActiveCell.Value, a_sReplace)
        Set funcFindCellRegExp = ActiveCell
    End If
End Function

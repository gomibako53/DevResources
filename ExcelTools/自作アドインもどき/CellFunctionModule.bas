Attribute VB_Name = "CellFunctionModule"
Option Explicit

' -------------------------------
' ����Function�𒼐ڃ��[�N�V�[�g�̐�������Ăяo�����Ƃ͂ł��Ȃ��̂ŁA
' ���̃��W���[�������[�N�V�[�g���̃��W���[���ɃR�s�[���Ďg�p����
' -------------------------------

' �Z���̔w�i�F���擾����
Public Function getCellColor(rng As Range) As Long
    getCellColor = rng.Interior.color
End Function

' �w��͈͂̃Z���Ɏw��̔w�i�F�̃Z�������݂��邩����
Public Function isExistCellColor(rng As Range, colorIndex As Long) As Boolean
    Dim r As Range
    For Each r In rng
    If r.Interior.color = colorIndex Then
        isExistCellColor = True
        Exit Function
    End If
    Next r
    isExistCellColor = False
End Function

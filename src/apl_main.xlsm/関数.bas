Attribute VB_Name = "�֐�"
'*******************************************************
'�֐��t�@�C��
'*******************************************************
Public genrecolumn As Long
Public QuestionCountColumn As Long

'�萔
Public Const GRNRECELL As String = "�W������"
Public Const QUESTION_COUNT_CELL As String = "�o���"
Public Const DBSheet As String = "DB"

'���f�[�^
Type QuestionData
        
    id As Long '����ID
    English As String '�p��
    Japanese As String '���{��
        
End Type

''���������s�v�B���񐳓����Ƃ��[���ɂȂ邩��
''Sub Initialize()
'
'    '�ȍ~�̏����Ŏg�������擾
'    '�W��������̗񐔏�����
'    genrecolumn = Sheet2.Range(GRNRECELL).Column
'    '�o��񐔗�̏�����
'    QuestionCountColumn = Sheet2.Range(QUESTION_COUNT_CELL).Column
'
'    '��x�o���������ēx�o���Ȃ��悤�ɂ��邽�߂ɏo��񐔂�������
'    '�o��񐔗��������
'    Range(Cells(2, QuestionCountColumn), Cells(2, QuestionCountColumn).End(xlDown)).Value = 0
'
''End Sub


'�����̃W�������̐���Ԃ�
Function GetWordNum(genre As String) As Long
    Dim wordcount As Long
    Dim TargetColums As Long
    
    With ThisWorkbook.Worksheets(DBSheet)
        TargetColums = .Range(GRNRECELL).Column
        wordcount = WorksheetFunction.CountIf(.Columns(TargetColums), genre)
    End With
    
    GetWordNum = wordcount
    
End Function

'���̃f�[�^��Ԃ�
Function GetQuestion() As QuestionData
    
    
    
End Function



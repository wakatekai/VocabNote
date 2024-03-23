'*******************************************************
'�֐��t�@�C��
'*******************************************************
Public genrecolumn As Long
Public QuestionCountColumn As Long

'�萔
Public Const IDCELL As String = "����ID"
Public Const GRNRECELL As String = "�W������"
Public Const ENGLISHCELL As String = "�p��"
Public Const JAPANESECELL As String = "���{��"
Public Const QUESTION_COUNT_CELL As String = "�o���"
Public Const DBSheet As String = "DB"

Enum enumGenre
    FRUIT = 0
    ALL
End Enum

Type QestionData
    longDBNumber As Long
    strQestionWord As String
    strAnswerWord As String
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

'���̃f�[�^��Ԃ��i�W����������A����ID�A���̒P��A�����̒P���Ԃ��j
'���얢�m�F
Function GetQuestion(genre As String) As QuestionData
    Dim genrecount As Long
    Dim QuestionIDsub As Long
    Dim IDColumns As Long
    Dim genreColums As Long
    Dim QuestionWordColumns As Long
    Dim QuestionAnswerColumns As Long
    Dim QuestionID As Long
    Dim QuestionAnswer As String
    Dim QuestionWord As String
    Dim QuestionIDcount As Long
    Dim QuestionIDcountRow As Long
    Dim TargetColums As Long
    
    '���̃W�������̐����J�E���g
    genrecount = GetWordNum(genre)
    
    '�����_���ɂ����ڂ��𐶐����A���̖��f�[�^��Ԃ�
    QuestionIDsub = Int(genrecoun * Rnd + 1)
    QuestionIDcount = 1
    QuestionIDcountRow = 2 ' 1�s�ڂ̓^�C�g���s�Ȃ̂�2�s�ڂ���J�E���g
    With ThisWorkbook.Worksheets(DBSheet)
         IDColumns = .Range(IDCELL).Column '����ID��
         genreColums = .Range(GRNRECELL).Column '�W��������
         QuestionWordColumns = .Range(ENGLISHCELL).Column '�p���
         QuestionAnswerColumns = .Range(JAPANESECELL).Column '���{���
         Do While QuestionIDcount <= QuestionIDsub
            If .Cells(QuestionIDcountRow, TargetColums) = genre Then
                QuestionID = .Cells(QuestionIDcountRow, IDColumns)
                QuestionWord = .Cells(QuestionIDcountRow, QuestionWordColumns)
                QuestionAnswer = .Cells(QuestionIDcountRow, QuestionAnswerColumns)
                QuestionIDcount = QuestionIDcount + 1
            End If
            QuestionIDcountRow = QuestionIDcountRow + 1
         Loop
    End With
    
    GetQuestion.longDBNumber = QuestionID
    GetQuestion.strQestionWord = QuestionWord
    GetQuestion.strAnswerWord = QuestionAnswer
    
End Function

'DB�ɂ�����f�[�^�̒����烉���_���œ��{����擾
Function GetWordRandomly(enGenre As enumGenre) As String
    Dim sheetDB As Worksheet
    Dim randomRow As Long
    Dim selectedGenre As String
    Dim lastRow As Long
    
    Set sheetDB = ThisWorkbook.Worksheets(DBSheet)
    
    ' �e�[�u���̃f�[�^�������Ă���s�����擾
    lastRow = sheetDB.Cells(sheetDB.Rows.Count, "A").End(xlUp).Row
    
    ' �����_���ɍs��I��
    randomRow = Application.WorksheetFunction.RandBetween(2, lastRow)
    
    ' �I�����ꂽ�W���������m�F
    Dim genreCol As Long
    genreCol = sheetDB.Range(GRNRECELL).Column
    selectedGenre = sheetDB.Cells(randomRow, genreCol)
    
    If selectedGenre = GetGenreName(enGenre) Or enGenre = ALL Then
        Dim japaneseCol As Long
        japaneseCol = sheetDB.Range(JAPANESECELL).Column
        GetWordRandomly = sheetDB.Cells(randomRow, japaneseCol)
    Else
        ' �W����������v���Ȃ��ꍇ�͍ċA�I�Ɋ֐����Ăяo��
        GetWordRandomly = GetWordRandomly(enGenre)
    End If
End Function

' �W�������ԍ��ɑΉ����閼�O���擾
Function GetGenreName(enGenre As enumGenre) As String
    Dim japaneseWord As String
    
    Select Case enGenre
        Case FRUIT
            japaneseWord = "�ʕ�"
        Case VEHICLE
            japaneseWord = "��蕨"
        Case ALL
            japaneseWord = "�S��"
    End Select
    
    GetGenreName = japaneseWord
End Function

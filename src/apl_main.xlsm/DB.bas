Attribute VB_Name = "DB"
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
Public Const CORRECTNUMCELL As String = "������"
Public Const DBSheet As String = "DB"

Enum enumGenre
    FRUIT = 0
    VEHICLE
    ALL
End Enum

Type QestionData
    longDBNumber As Long
    strQestionWord As String
    strAnswerWord As String
End Type

'�����̃W�������̐���Ԃ�
Function GetWordNum(genre As enumGenre) As Long
    Dim wordcount As Long
    Dim TargetColums As Long
    
    With ThisWorkbook.Worksheets(DBSheet)
        TargetColums = .Range(GRNRECELL).Column
        If genre = ALL Then
            '�^�C�g���s������̂�-1
            wordcount = .Range("A1").CurrentRegion.Rows.Count - 1
        Else
            wordcount = WorksheetFunction.CountIf(.Columns(TargetColums), genre)
        End If
    End With
    
    GetWordNum = wordcount
    
End Function

'���̃f�[�^��Ԃ��i�W����������A����ID�A���̒P��A�����̒P���Ԃ��j
Function GetQuestion(genre As enumGenre) As QuestionData
    Dim genrecount As Long
    Dim QuestionIDsub As Long  '�����̃W�������̏ォ�牽�Ԗڂ̖��f�[�^���擾���邩
    Dim IDColumns As Long
    Dim genreColums As Long
    Dim QuestionWordColumns As Long
    Dim QuestionAnswerColumns As Long
    Dim QuestionID As Long
    Dim QuestionAnswer As String
    Dim QuestionWord As String
    Dim QuestionIDcount As Long
    Dim QuestionIDcountRow As Long
    
    '���̃W�������̐����J�E���g
    genrecount = GetWordNum(genre)
    
    '�����_���ɂ����ڂ��𐶐����A���̖��f�[�^��Ԃ�
    QuestionIDsub = Int(genrecount * Rnd + 1)
    QuestionIDcount = 1
    QuestionIDcountRow = 2 ' 1�s�ڂ̓^�C�g���s�Ȃ̂�2�s�ڂ���J�E���g
    With ThisWorkbook.Worksheets(DBSheet)
         IDColumns = .Range(IDCELL).Column '����ID��
         genreColums = .Range(GRNRECELL).Column '�W��������
         QuestionWordColumns = .Range(ENGLISHCELL).Column '�p���
         QuestionAnswerColumns = .Range(JAPANESECELL).Column '���{���
         If genre = ALL Then
            Do While QuestionIDcount <= QuestionIDsub
                QuestionID = .Cells(QuestionIDcountRow, IDColumns)
                QuestionWord = .Cells(QuestionIDcountRow, QuestionWordColumns)
                QuestionAnswer = .Cells(QuestionIDcountRow, QuestionAnswerColumns)
                QuestionIDcount = QuestionIDcount + 1
                QuestionIDcountRow = QuestionIDcountRow + 1
            Loop
         Else
            Do While QuestionIDcount <= QuestionIDsub
                If .Cells(QuestionIDcountRow, genreColums) = genre Then
                    QuestionID = .Cells(QuestionIDcountRow, IDColumns)
                    QuestionWord = .Cells(QuestionIDcountRow, QuestionWordColumns)
                    QuestionAnswer = .Cells(QuestionIDcountRow, QuestionAnswerColumns)
                    QuestionIDcount = QuestionIDcount + 1
                End If
                QuestionIDcountRow = QuestionIDcountRow + 1
            Loop
        End If
    End With
    
    GetQuestion.longDBNumber = QuestionID
    GetQuestion.strQestionWord = QuestionWord
    GetQuestion.strAnswerWord = QuestionAnswer
End Function

'DB�ɂ�����f�[�^�̒����烉���_���œ��{����擾
Function GetWrongWord(enGenre As enumGenre) As String
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
    
    If selectedGenre = enGenre Or enGenre = ALL Then
        Dim japaneseCol As Long
        japaneseCol = sheetDB.Range(JAPANESECELL).Column
        GetWrongWord = sheetDB.Cells(randomRow, japaneseCol)
    Else
        ' �W����������v���Ȃ��ꍇ�͍ċA�I�Ɋ֐����Ăяo��
        GetWrongWord = GetWrongWord(enGenre)
    End If
End Function

'����ʒm (�񓚂̐������DB���X�V)
Private Sub SetAnswer(DBNumber As Long, blResult As Boolean)
    '�C���N�������g�̏���K�[�h�͉ǐ���D�悵�Ė���
    Dim FindIDRange As Range
    
    Set FindIDRange = Range(Range(IDCELL).Cells, Range(IDCELL).End(xlDown).Cells).Find(CStr(DBNumber), LookAt:=xlWhole)
    If FindIDRange Is Nothing Then
        MsgBox "�w�肳�ꂽ����ID�͑��݂��܂���"
        Exit Sub
    Else
        '�o��񐔃C���N�������g
        Cells(FindIDRange.Row, Range(QUESTION_COUNT_CELL).Column).Value = Cells(FindIDRange.Row, Range(QUESTION_COUNT_CELL).Column).Value + 1
        
        If blResult Then
            '�����񐔃C���N�������g
            Cells(FindIDRange.Row, Range(CORRECTNUMCELL).Column).Value = Cells(FindIDRange.Row, Range(CORRECTNUMCELL).Column).Value + 1
        End If
    End If

End Sub


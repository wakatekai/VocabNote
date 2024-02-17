Attribute VB_Name = "apl_main"
Option Explicit

Const WLONG_WORD_NUM As Long = 3 - 1
Const QUESTION_NUM As Long = 5

'���f�[�^
Type QestionData
    longDBNumber As Long
    strQestionWord As String
    strAnswerWord As String
    strWrongWord(WLONG_WORD_NUM) As String
End Type

'�W������
Enum enumGenre
    FRUIT = 0
    ALL
End Enum

Sub apl_main()
    Dim stQestionData As QestionData '���[�U�[�C���^�[�t�F�[�X���ŃO���[�o���ϐ��ɂ��Ȃ��Ǝg�p�ł��Ȃ��悤�Ȃ�ύX�i��������ƈ����������ݕs�v�ɂȂ�A���o�͂̈Ӗ��������c�H�j
    Dim blEndFlag As Boolean
    Dim enGenre As enumGenre
    Dim blResult As Boolean
    Dim longNumQuestions As Long
    Dim longNumCorrectAnswers As Long
    Dim i As Long
    Dim j As Long
    Dim blDuplicate As Boolean
    
    
    '���^�C�g���\����
    blEndFlag = DispTitle()
    '���I�����聄
    '�I���t���O����������֐����I��
    If blEndFlag = True Then
        Exit Sub
    End If
    
    enGenre = FRUIT '�b��I�ɃW�������͌Œ�i�����I�ɂ̓^�C�g���őI���ł���Ƃ��������j
    
    longNumQuestions = 0 '�񓚐�������
    longNumCorrectAnswers = 0 '���𐔏�����
    
    For longNumQuestions = 0 To QUESTION_NUM Step 1
        '�����f�[�^�擾��
        '�Q�Ɠn���ɂ��ăR�[�������֐����ŕϐ���ύX���Ă��炤�C���[�W�@�\���̕ϐ����ۂ��ƍs������������͂悳����
        Call GetQuestion(enGenre, stQestionData.longDBNumber, stQestionData.strQestionWord, stQestionData.strAnswerWord)
        
        '���듚�f�[�^�擾��
        '1�ꂸ�擾
        i = 0
        While i <= WLONG_WORD_NUM
            stQestionData.strWrongWord(i) = GetWrongWord(enGenre)
            '�d���m�F
            blDuplicate = CheckDuplicates(stQestionData.strWrongWord)
            If blDuplicate = False Then
                i = i + 1
            End If
        Wend
        
        
        '�����\���E���ʔ��聄
        blResult = SetQuestion(stQestionData.strQestionWord, stQestionData.strAnswerWord, stQestionData.strWrongWord)
        If blResult = True Then
            longNumCorrectAnswers = longNumCorrectAnswers + 1 '���𐔃C���N�������g
        End If
        
        '������ʒm��
        Call SetAnswer(stQestionData.longDBNumber, blResult)

    Next longNumQuestions
    
    '�����ʕ\����
    Call DispResult(longNumQuestions, longNumCorrectAnswers)
    
End Sub

Function CheckDuplicates(arr() As String) As Boolean
    Dim i As Long, j As Long
    
    For i = LBound(arr) To UBound(arr)
        For j = i + 1 To UBound(arr)
            If arr(i) <> "" And arr(j) <> "" And arr(i) = arr(j) Then
                CheckDuplicates = True
                Exit Function
            End If
        Next j
    Next i
    
    CheckDuplicates = False
End Function

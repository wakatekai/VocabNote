Attribute VB_Name = "apl_main"
'*******************************************************
'�֐��t�@�C��
'*******************************************************

Option Explicit

Const WLONG_WORD_NUM As Long = 3
Const QUESTION_NUM As Long = 5
Const CHOICIES_NUM As Long = 4

'���f�[�^
Type QestionData
    longDBNumber As Long
    strQestionWord As String
    strAnswerWord As String
    strWrongWord(WLONG_WORD_NUM) As String
End Type


Sub apl_main()
    Dim stQestionData As QestionData
    Dim blEndFlag As Boolean
    Dim enGenre As enumGenre
    Dim blResult As Boolean
    Dim longNumQuestions As Long
    Dim longNumCorrectAnswers As Long
    Dim i As Long
    Dim j As Long
    Dim blDuplicate As Boolean
    Dim strChoices(CHOICIES_NUM) As String
    
    
    Do
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
            '4/13 DB���œƎ��̍\���̂��`���A�߂�l�ɐݒ肵�Ă��邽�߁A�f�[�^���s���������A�R���p�C���G���[
            '�\���̕ϐ���߂�l�ɂ���Ƃǂ̃f�[�^���ݒ肳�ꂽ���킩��Ȃ����ߌ㏈���Őݒ肳��Ă��Ȃ��f�[�^���Q�Ƃ��Ȃ����߂̏��u���K�v
            Call GetQuestion(enGenre, stQestionData.longDBNumber, stQestionData.strQestionWord, stQestionData.strAnswerWord)
            
            '���듚�f�[�^�擾��
            '1�ꂸ�擾
            strChoices(0) = stQestionData.strAnswerWord '�d���m�F�p�ɑI�����z��擪�ɓ��������Ă���
            i = 0
            While i <= WLONG_WORD_NUM
                strChoices(i + 1) = GetWrongWord(enGenre)   '�I�����z��Ɍ듚�����Ă���
                '�d���m�F
                blDuplicate = CheckDuplicates(strChoices)
                If blDuplicate = False Then
                    stQestionData.strWrongWord(i) = strChoices(i + 1)  '���̑I�����Ɣ��Ȃ��������ߌ듚�Ƃ��ēo�^
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
    Loop While DispResult(longNumQuestions, longNumCorrectAnswers)
    
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

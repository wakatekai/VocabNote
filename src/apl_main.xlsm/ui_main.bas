Attribute VB_Name = "ui_main"
'���ʎg�p�̕ϐ��i�f�o�b�O�p�j
Const WLONG_WORD_NUM As Long = 3 - 1

Type QuestionData
 longDBNumber As Long
 strQuestionWord As String
 strAnswerWord As String
 strWrongWord(WLONG_WORD_NUM) As String
End Type


'UI�ŗL�̕ϐ�
Dim ToF As Long           '���딻��t���O
Dim UiblEndFlag As Long   '���J�n����t���O



Public Ans As Integer    '���딻��
'���[�h����֐����쐬����
Function InterFace(strQestionWord As String, strAnswerWord As String, strWrongWord0 As String, strWrongWord1 As String, strWrongWord2 As String)
    Dim UI As QuestionData
'    strAnswerWord = "anwer"
'    strWrongWord0 = "wrong1"
'    strWrongWord1 = "wrong2"
'    strWrongWord2 = "wrong3"
    
    '�o��p�^�[���̐���
    Dim intMax As Integer '�ő�l
    Dim intMin As Integer '�ŏ��l
    Dim Ptrn As Integer   '�o��p�^�[��
    
    intMax = 3
    intMin = 1
    
    Ptrn = Int((intMax - intMin + 1) * Rnd + intMin)
    Select Case Ptrn
        Case 1
            Answer.Label1.Caption = strAnswerWord
            Answer.Label2.Caption = strWrongWord0
            Answer.Label3.Caption = strWrongWord1
            Answer.Label4.Caption = strWrongWord2
            Answer.Label6.Caption = strQestionWord
            Ans = 1
        Case 2
            Answer.Label1.Caption = strWrongWord0
            Answer.Label2.Caption = strAnswerWord
            Answer.Label3.Caption = strWrongWord1
            Answer.Label4.Caption = strWrongWord2
            Answer.Label6.Caption = strQestionWord
            Ans = 2
        Case 3
            Answer.Label1.Caption = strWrongWord0
            Answer.Label2.Caption = strWrongWord1
            Answer.Label3.Caption = strAnswerWord
            Answer.Label4.Caption = strWrongWord0
            Answer.Label6.Caption = strQestionWord
            Ans = 3
        Case Else
            Answer.Label1.Caption = strWrongWord0
            Answer.Label2.Caption = strWrongWord1
            Answer.Label3.Caption = strWrongWord2
            Answer.Label4.Caption = strAnswerWord
            Answer.Label6.Caption = strQestionWord
            Ans = 4
    End Select

    '�o���ʌĂяo��
    Answer.Show     '���\��/���ʔ����ʂ̌Ăяo��
End Function

'���딻��̊֐�
Function Func1(ByVal Ans As String, ByVal SlctNm As String) As String
    If Ans = SlctNm Then
        Func1 = "�Z"
        ToF = True
    Else
        Func1 = "�~"
        ToF = False
    End If
End Function

Function DispTitle()
'    Dim UiblEndFlag             '���J�n����t���O
    UiblEndFlag = 4
    Title.Show              '�^�C�g����ʂ̌Ăяo��
    DispTitle = UiblEndFlag
End Function

Sub SorE(EndFlag As Long)
    If EndFlag = True Then
        UiblEndFlag = True      '�A�v���I��
    Else
        UiblEndFlag = False     '���J�n
    End If
End Sub


Function SetQuestion(strQestionWord As String, strAnswerWord As String, strWrongWord() As String)
'    Dim ToF As Long           '���딻��t���O
    Call InterFace(strQestionWord, strAnswerWord, strWrongWord(0), strWrongWord(1), strWrongWord(2))
    SetQuestion = ToF
'    Answer.Show  ���\��/���ʔ����InterFace()�ŌĂяo��
End Function

Function DispResult(longNumQuestions As Long, longNumCorrectAnswers As Long)
    Result.Label1.Caption = longNumQuestions
    Result.Label2.Caption = longNumCorrectAnswers
    Result.Show       '���ʔ����ʂ̌Ăяo��
    DispResult = True '�^�C�g����ʂ֖߂邽��True��Ԃ�
End Function

Attribute VB_Name = "Module11"
'���ʎg�p�̕ϐ��i�f�o�b�O�p�j
Const WLONG_WORD_NUM As Long = 3 - 1

Type QuestionData
 longDBNumber As Long
 strQuestionWord As String
 strAnswerWord As String
 strWrongWord(WLONG_WORD_NUM) As String
End Type


'UI�ŗL�̕ϐ�
Dim ToF As Long                 '���딻��t���O
Public UiblEndFlag As Boolean   '���J�n����t���O

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
            UserForm4.Label1.Caption = strAnswerWord
            UserForm4.Label2.Caption = strWrongWord0
            UserForm4.Label3.Caption = strWrongWord1
            UserForm4.Label4.Caption = strWrongWord2
            UserForm4.Label6.Caption = strQestionWord
            Ans = 1
        Case 2
            UserForm4.Label1.Caption = strWrongWord0
            UserForm4.Label2.Caption = strAnswerWord
            UserForm4.Label3.Caption = strWrongWord1
            UserForm4.Label4.Caption = strWrongWord2
            UserForm4.Label6.Caption = strQestionWord
            Ans = 2
        Case 3
            UserForm4.Label1.Caption = strWrongWord0
            UserForm4.Label2.Caption = strWrongWord1
            UserForm4.Label3.Caption = strAnswerWord
            UserForm4.Label4.Caption = strWrongWord0
            UserForm4.Label6.Caption = strQestionWord
            Ans = 3
        Case Else
            UserForm4.Label1.Caption = strWrongWord0
            UserForm4.Label2.Caption = strWrongWord1
            UserForm4.Label3.Caption = strWrongWord2
            UserForm4.Label4.Caption = strAnswerWord
            UserForm4.Label6.Caption = strQestionWord
            Ans = 4
    End Select

    '�o���ʌĂяo��
    UserForm4.Show     '���\��/���ʔ����ʂ̌Ăяo��
End Function

'���딻��̊֐�
Function Func1(ByVal Ans As String, ByVal SlctNm As String) As String
    If Ans = SlctNm Then
        Func1 = "�Z"
    Else
        Func1 = "�~"
    End If
End Function

Function DispTitle() As Boolean
    UiblEndFlag = True
    Title.Show              '�^�C�g����ʂ̌Ăяo��
    DispTitle = UiblEndFlag
End Function

Function SetQuestion(strQestionWord As String, strAnswerWord As String, strWrongWord() As String)
'    Dim ToF As Long           '���딻��t���O
    Call InterFace(strQestionWord, strAnswerWord, strWrongWord(0), strWrongWord(1), strWrongWord(2))
    SetQuestion = ToF
'    UserForm4.Show  ���\��/���ʔ����InterFace()�ŌĂяo��
End Function

Sub DispResult(longNumQuestions As Long, longNumCorrectAnswers As Long)
    'Call InterFace(
    UserForm8.Show       '���ʔ����ʂ̌Ăяo��
End Sub



Attribute VB_Name = "ui_main"
'共通使用の変数（デバッグ用）
Const WLONG_WORD_NUM As Long = 3 - 1

Type QuestionData
 longDBNumber As Long
 strQuestionWord As String
 strAnswerWord As String
 strWrongWord(WLONG_WORD_NUM) As String
End Type


'UI固有の変数
Dim ToF As Long           '正誤判定フラグ
Dim UiblEndFlag As Long   '問題開始判定フラグ



Public Ans As Integer    '正誤判定
'ロードする関数を作成する
Function InterFace(strQestionWord As String, strAnswerWord As String, strWrongWord0 As String, strWrongWord1 As String, strWrongWord2 As String)
    Dim UI As QuestionData
'    strAnswerWord = "anwer"
'    strWrongWord0 = "wrong1"
'    strWrongWord1 = "wrong2"
'    strWrongWord2 = "wrong3"
    
    '出題パターンの生成
    Dim intMax As Integer '最大値
    Dim intMin As Integer '最小値
    Dim Ptrn As Integer   '出題パターン
    
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

    '出題画面呼び出し
    Answer.Show     '問題表示/結果判定画面の呼び出し
End Function

'正誤判定の関数
Function Func1(ByVal Ans As String, ByVal SlctNm As String) As String
    If Ans = SlctNm Then
        Func1 = "〇"
        ToF = True
    Else
        Func1 = "×"
        ToF = False
    End If
End Function

Function DispTitle()
'    Dim UiblEndFlag             '問題開始判定フラグ
    UiblEndFlag = 4
    Title.Show              'タイトル画面の呼び出し
    DispTitle = UiblEndFlag
End Function

Sub SorE(EndFlag As Long)
    If EndFlag = True Then
        UiblEndFlag = True      'アプリ終了
    Else
        UiblEndFlag = False     '問題開始
    End If
End Sub


Function SetQuestion(strQestionWord As String, strAnswerWord As String, strWrongWord() As String)
'    Dim ToF As Long           '正誤判定フラグ
    Call InterFace(strQestionWord, strAnswerWord, strWrongWord(0), strWrongWord(1), strWrongWord(2))
    SetQuestion = ToF
'    Answer.Show  問題表示/結果判定はInterFace()で呼び出す
End Function

Function DispResult(longNumQuestions As Long, longNumCorrectAnswers As Long)
    Result.Label1.Caption = longNumQuestions
    Result.Label2.Caption = longNumCorrectAnswers
    Result.Show       '結果判定画面の呼び出し
    DispResult = True 'タイトル画面へ戻るためTrueを返す
End Function

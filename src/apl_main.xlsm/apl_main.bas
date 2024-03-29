Attribute VB_Name = "apl_main"
Option Explicit

Const WLONG_WORD_NUM As Long = 3 - 1
Const QUESTION_NUM As Long = 5

'問題データ
Type QestionData
    longDBNumber As Long
    strQestionWord As String
    strAnswerWord As String
    strWrongWord(WLONG_WORD_NUM) As String
End Type

'ジャンル
Enum enumGenre
    FRUIT = 0
    ALL
End Enum

Sub apl_main()
    Dim stQestionData As QestionData 'ユーザーインターフェース側でグローバル変数にしないと使用できないようなら変更（そうすると引数が軒並み不要になり、入出力の意味が薄れる…？）
    Dim blEndFlag As Boolean
    Dim enGenre As enumGenre
    Dim blResult As Boolean
    Dim longNumQuestions As Long
    Dim longNumCorrectAnswers As Long
    Dim i As Long
    Dim j As Long
    Dim blDuplicate As Boolean
    
    Do
    
        '＜タイトル表示＞
        blEndFlag = DispTitle()
        '＜終了判定＞
        '終了フラグが立ったら関数を終了
        If blEndFlag = True Then
            Exit Sub
        End If
        
        enGenre = FRUIT '暫定的にジャンルは固定（将来的にはタイトルで選択できるといいかも）
        
        longNumQuestions = 0 '回答数初期化
        longNumCorrectAnswers = 0 '正解数初期化
        
        For longNumQuestions = 0 To QUESTION_NUM Step 1
            '＜問題データ取得＞
            '参照渡しにしてコールした関数側で変数を変更してもらうイメージ　構造体変数を丸ごと行き来させるよりはよさそう
            Call GetQuestion(enGenre, stQestionData.longDBNumber, stQestionData.strQestionWord, stQestionData.strAnswerWord)
            
            '＜誤答データ取得＞
            '1語ずつ取得
            i = 0
            While i <= WLONG_WORD_NUM
                stQestionData.strWrongWord(i) = GetWrongWord(enGenre)
                '重複確認
                blDuplicate = CheckDuplicates(stQestionData.strWrongWord)
                If blDuplicate = False Then
                    i = i + 1
                End If
            Wend
            
            
            '＜問題表示・結果判定＞
            blResult = SetQuestion(stQestionData.strQestionWord, stQestionData.strAnswerWord, stQestionData.strWrongWord)
            If blResult = True Then
                longNumCorrectAnswers = longNumCorrectAnswers + 1 '正解数インクリメント
            End If
            
            '＜正誤通知＞
            Call SetAnswer(stQestionData.longDBNumber, blResult)
    
        Next longNumQuestions
        
        '＜結果表示＞
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

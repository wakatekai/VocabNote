Attribute VB_Name = "apl_main"
'*******************************************************
'関数ファイル
'*******************************************************

Option Explicit

Const WLONG_WORD_NUM As Long = 3
Const QUESTION_NUM As Long = 5
Const CHOICIES_NUM As Long = 4

'問題データ
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
            '4/13 DB側で独自の構造体を定義し、戻り値に設定しているため、データが行き来せず、コンパイルエラー
            '構造体変数を戻り値にするとどのデータが設定されたかわからないため後処理で設定されていないデータを参照しないための処置が必要
            Call GetQuestion(enGenre, stQestionData.longDBNumber, stQestionData.strQestionWord, stQestionData.strAnswerWord)
            
            '＜誤答データ取得＞
            '1語ずつ取得
            strChoices(0) = stQestionData.strAnswerWord '重複確認用に選択肢配列先頭に答えを入れておく
            i = 0
            While i <= WLONG_WORD_NUM
                strChoices(i + 1) = GetWrongWord(enGenre)   '選択肢配列に誤答を入れておく
                '重複確認
                blDuplicate = CheckDuplicates(strChoices)
                If blDuplicate = False Then
                    stQestionData.strWrongWord(i) = strChoices(i + 1)  '他の選択肢と被らなかったため誤答として登録
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

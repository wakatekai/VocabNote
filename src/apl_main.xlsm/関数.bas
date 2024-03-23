'*******************************************************
'関数ファイル
'*******************************************************
Public genrecolumn As Long
Public QuestionCountColumn As Long

'定数
Public Const IDCELL As String = "識別ID"
Public Const GRNRECELL As String = "ジャンル"
Public Const ENGLISHCELL As String = "英語"
Public Const JAPANESECELL As String = "日本語"
Public Const QUESTION_COUNT_CELL As String = "出題回数"
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

''初期化←不要。毎回正答率とかゼロになるから
''Sub Initialize()
'
'    '以降の処理で使う情報を取得
'    'ジャンル列の列数初期化
'    genrecolumn = Sheet2.Range(GRNRECELL).Column
'    '出題回数列の初期化
'    QuestionCountColumn = Sheet2.Range(QUESTION_COUNT_CELL).Column
'
'    '一度出した問題を再度出さないようにするために出題回数を初期化
'    '出題回数列を初期化
'    Range(Cells(2, QuestionCountColumn), Cells(2, QuestionCountColumn).End(xlDown)).Value = 0
'
''End Sub


'引数のジャンルの数を返す
Function GetWordNum(genre As String) As Long
    Dim wordcount As Long
    Dim TargetColums As Long
    
    With ThisWorkbook.Worksheets(DBSheet)
        TargetColums = .Range(GRNRECELL).Column
        wordcount = WorksheetFunction.CountIf(.Columns(TargetColums), genre)
    End With
    
    GetWordNum = wordcount
    
End Function

'問題のデータを返す（ジャンルから、識別ID、問題の単語、答えの単語を返す）
'動作未確認
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
    
    'そのジャンルの数をカウント
    genrecount = GetWordNum(genre)
    
    'ランダムにいくつ目かを生成し、その問題データを返す
    QuestionIDsub = Int(genrecoun * Rnd + 1)
    QuestionIDcount = 1
    QuestionIDcountRow = 2 ' 1行目はタイトル行なので2行目からカウント
    With ThisWorkbook.Worksheets(DBSheet)
         IDColumns = .Range(IDCELL).Column '識別ID列
         genreColums = .Range(GRNRECELL).Column 'ジャンル列
         QuestionWordColumns = .Range(ENGLISHCELL).Column '英語列
         QuestionAnswerColumns = .Range(JAPANESECELL).Column '日本語列
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

'DBにある問題データの中からランダムで日本語を取得
Function GetWordRandomly(enGenre As enumGenre) As String
    Dim sheetDB As Worksheet
    Dim randomRow As Long
    Dim selectedGenre As String
    Dim lastRow As Long
    
    Set sheetDB = ThisWorkbook.Worksheets(DBSheet)
    
    ' テーブルのデータが入っている行数を取得
    lastRow = sheetDB.Cells(sheetDB.Rows.Count, "A").End(xlUp).Row
    
    ' ランダムに行を選択
    randomRow = Application.WorksheetFunction.RandBetween(2, lastRow)
    
    ' 選択されたジャンルを確認
    Dim genreCol As Long
    genreCol = sheetDB.Range(GRNRECELL).Column
    selectedGenre = sheetDB.Cells(randomRow, genreCol)
    
    If selectedGenre = GetGenreName(enGenre) Or enGenre = ALL Then
        Dim japaneseCol As Long
        japaneseCol = sheetDB.Range(JAPANESECELL).Column
        GetWordRandomly = sheetDB.Cells(randomRow, japaneseCol)
    Else
        ' ジャンルが一致しない場合は再帰的に関数を呼び出す
        GetWordRandomly = GetWordRandomly(enGenre)
    End If
End Function

' ジャンル番号に対応する名前を取得
Function GetGenreName(enGenre As enumGenre) As String
    Dim japaneseWord As String
    
    Select Case enGenre
        Case FRUIT
            japaneseWord = "果物"
        Case VEHICLE
            japaneseWord = "乗り物"
        Case ALL
            japaneseWord = "全部"
    End Select
    
    GetGenreName = japaneseWord
End Function

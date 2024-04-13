Attribute VB_Name = "DB"
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
Public Const CORRECTNUMCELL As String = "正解回数"
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

'引数のジャンルの数を返す
Function GetWordNum(genre As enumGenre) As Long
    Dim wordcount As Long
    Dim TargetColums As Long
    
    With ThisWorkbook.Worksheets(DBSheet)
        TargetColums = .Range(GRNRECELL).Column
        If genre = ALL Then
            'タイトル行があるので-1
            wordcount = .Range("A1").CurrentRegion.Rows.Count - 1
        Else
            wordcount = WorksheetFunction.CountIf(.Columns(TargetColums), genre)
        End If
    End With
    
    GetWordNum = wordcount
    
End Function

'問題のデータを返す（ジャンルから、識別ID、問題の単語、答えの単語を返す）
Function GetQuestion(genre As enumGenre) As QuestionData
    Dim genrecount As Long
    Dim QuestionIDsub As Long  '引数のジャンルの上から何番目の問題データを取得するか
    Dim IDColumns As Long
    Dim genreColums As Long
    Dim QuestionWordColumns As Long
    Dim QuestionAnswerColumns As Long
    Dim QuestionID As Long
    Dim QuestionAnswer As String
    Dim QuestionWord As String
    Dim QuestionIDcount As Long
    Dim QuestionIDcountRow As Long
    
    'そのジャンルの数をカウント
    genrecount = GetWordNum(genre)
    
    'ランダムにいくつ目かを生成し、その問題データを返す
    QuestionIDsub = Int(genrecount * Rnd + 1)
    QuestionIDcount = 1
    QuestionIDcountRow = 2 ' 1行目はタイトル行なので2行目からカウント
    With ThisWorkbook.Worksheets(DBSheet)
         IDColumns = .Range(IDCELL).Column '識別ID列
         genreColums = .Range(GRNRECELL).Column 'ジャンル列
         QuestionWordColumns = .Range(ENGLISHCELL).Column '英語列
         QuestionAnswerColumns = .Range(JAPANESECELL).Column '日本語列
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

'DBにある問題データの中からランダムで日本語を取得
Function GetWrongWord(enGenre As enumGenre) As String
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
    
    If selectedGenre = enGenre Or enGenre = ALL Then
        Dim japaneseCol As Long
        japaneseCol = sheetDB.Range(JAPANESECELL).Column
        GetWrongWord = sheetDB.Cells(randomRow, japaneseCol)
    Else
        ' ジャンルが一致しない場合は再帰的に関数を呼び出す
        GetWrongWord = GetWrongWord(enGenre)
    End If
End Function

'正誤通知 (回答の正誤情報でDBを更新)
Private Sub SetAnswer(DBNumber As Long, blResult As Boolean)
    'インクリメントの上限ガードは可読性を優先して無し
    Dim FindIDRange As Range
    
    Set FindIDRange = Range(Range(IDCELL).Cells, Range(IDCELL).End(xlDown).Cells).Find(CStr(DBNumber), LookAt:=xlWhole)
    If FindIDRange Is Nothing Then
        MsgBox "指定された識別IDは存在しません｡"
        Exit Sub
    Else
        '出題回数インクリメント
        Cells(FindIDRange.Row, Range(QUESTION_COUNT_CELL).Column).Value = Cells(FindIDRange.Row, Range(QUESTION_COUNT_CELL).Column).Value + 1
        
        If blResult Then
            '正解回数インクリメント
            Cells(FindIDRange.Row, Range(CORRECTNUMCELL).Column).Value = Cells(FindIDRange.Row, Range(CORRECTNUMCELL).Column).Value + 1
        End If
    End If

End Sub


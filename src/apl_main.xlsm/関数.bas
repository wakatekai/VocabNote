Attribute VB_Name = "関数"
'*******************************************************
'関数ファイル
'*******************************************************
Public genrecolumn As Long
Public QuestionCountColumn As Long

'定数
Public Const GRNRECELL As String = "ジャンル"
Public Const QUESTION_COUNT_CELL As String = "出題回数"
Public Const DBSheet As String = "DB"

'問題データ
Type QuestionData
        
    id As Long '識別ID
    English As String '英語
    Japanese As String '日本語
        
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

'問題のデータを返す
Function GetQuestion() As QuestionData
    
    
    
End Function



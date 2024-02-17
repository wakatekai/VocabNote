Attribute VB_Name = "DB_stub"
Option Explicit

Public Sub GetQuestion(enGenre As enumGenre, longDBNumber As Long, strQestionWord As String, strAnswerWord As String)
    Dim Genre As Long
    Genre = enGenre
    longDBNumber = 1
    strQestionWord = "Question"
    strAnswerWord = "Answer"
End Sub

Public Function GetWrongWord(enGenre As enumGenre)
    Dim Genre As Long
    
    Genre = enGenre
    GetWrongWord = "Wrong"
    
End Function

Public Sub SetAnswer(longDBNumber As Long, blResult As Boolean)
    Dim Genre As Long
    Dim DBNumber As Long
    Dim Result As Boolean
    
    DBNumber = longDBNumber
    Result = blResult
    
End Sub


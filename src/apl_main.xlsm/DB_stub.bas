Attribute VB_Name = "DB_stub"
Option Explicit

Dim Number As Long

Public Sub GetQuestion(enGenre As enumGenre, longDBNumber As Long, strQestionWord As String, strAnswerWord As String)
    Dim Genre As Long
    Genre = enGenre
    longDBNumber = 1
    strQestionWord = "Question"
    strAnswerWord = "Answer"
End Sub

Public Function GetWrongWord(enGenre As enumGenre)
    Dim Genre As Long

    Number = Number + 1

    Genre = enGenre
    GetWrongWord = "Wrong" & Number
    
End Function

Public Sub SetAnswer(longDBNumber As Long, blResult As Boolean)
    Dim Genre As Long
    Dim DBNumber As Long
    Dim result As Boolean
    
    DBNumber = longDBNumber
    result = blResult
    
End Sub


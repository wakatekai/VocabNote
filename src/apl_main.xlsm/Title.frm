VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Title 
   Caption         =   "タイトル"
   ClientHeight    =   4557
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7560
   OleObjectBlob   =   "Title.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndFlag As Long '問題開始:False/アプリ終了:True

Option Explicit
Private Sub CB_START_Click()
    '[START]をクリック
    Unload Title
    SorE (EndFlag)
End Sub

Private Sub CB_END_Click()
    '[END]をクリック
    Unload Title
    EndFlag = True
    SorE (EndFlag)
End Sub

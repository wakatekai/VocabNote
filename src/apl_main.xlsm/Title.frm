VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Title 
   Caption         =   "�^�C�g��"
   ClientHeight    =   4560
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   7560
   OleObjectBlob   =   "Title.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CB_START_Click()
    '[START]���N���b�N
    UiblEndFlag = False
    Unload Title
End Sub

Private Sub CB_END_Click()
    '[END]���N���b�N
    UiblEndFlag = True
    Unload Title
End Sub



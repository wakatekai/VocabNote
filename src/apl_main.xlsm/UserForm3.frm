VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4110
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   7395
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
'���J�n���N���b�N
Unload UserForm3
UiblEndFlag = False
'InterFace
'UserForm4.Show
End Sub

Private Sub CommandButton2_Click()
'�A�v���I�����N���b�N
Unload UserForm3
UiblEndFlag = True
End Sub

Private Sub UserForm_Click()

End Sub

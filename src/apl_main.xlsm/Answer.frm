VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Answer 
   Caption         =   "��"
   ClientHeight    =   6328
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   11312
   OleObjectBlob   =   "Answer.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Answer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTALCOUNT
Dim SlctNm As Integer


Private Sub CommandButton1_Click() 'A�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 1
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
End If
End Sub

Private Sub CommandButton2_Click() 'B�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 2
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
End If
End Sub

Private Sub CommandButton3_Click() 'C�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 3
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
End If
End Sub

Private Sub CommandButton4_Click() 'D�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 4
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
End If
End Sub

Private Sub CommandButton5_Click()  '���̖��փ{�^��
If TOTALCOUNT = 0 Then
    MsgBox "�X�L�b�v���܂�����낵���ł����H"
    TOTALCOUNT = 2
ElseIf TOTALCOUNT = 1 Or TOTALCOUNT = 2 Then
    Unload Answer
'    Result.Show
End If
'MsgBox "�X�L�b�v���܂�����낵���ł����H"
End Sub


Private Sub CommandButton6_Click()  '�^�C�g����ʂփ{�^��
'���g�p
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub test()
    MsgBox "test"
End Sub


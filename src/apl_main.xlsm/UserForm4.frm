VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   6330
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   11310
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm4"
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
    ToF = Func1(Ans, SlctNm)             '���딻�茋�ʎ擾
End If
End Sub

Private Sub CommandButton2_Click() 'B�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 2
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
    ToF = Func1(Ans, SlctNm)             '���딻�茋�ʎ擾
End If
End Sub

Private Sub CommandButton3_Click() 'C�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 3
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
    ToF = Func1(Ans, SlctNm)             '���딻�茋�ʎ擾
End If
End Sub

Private Sub CommandButton4_Click() 'D�{�^��
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 4
    Label5.Caption = Func1(Ans, SlctNm)  '���딻�茋�ʂ̉�ʕ\��
    ToF = Func1(Ans, SlctNm)             '���딻�茋�ʎ擾
End If
End Sub

Private Sub CommandButton5_Click()  '���̖��փ{�^��
If TOTALCOUNT = 0 Then
    MsgBox "�X�L�b�v���܂�����낵���ł����H"
    TOTALCOUNT = 2
ElseIf TOTALCOUNT = 1 Or TOTALCOUNT = 2 Then
    Unload UserForm4
'    UserForm8.Show
End If
'MsgBox "�X�L�b�v���܂�����낵���ł����H"
End Sub


Private Sub CommandButton6_Click()  '�^�C�g����ʂփ{�^��
'If TOTALCOUNT = 3 Then
'    Unload UserForm4
'    UserForm3.Show
'End If
'If TOTALCOUNT = 0 Or 1 Then
'    MsgBox "�^�C�g����ʂ֖߂�܂�����낵���ł����H"
'    TOTALCOUNT = 3
'End If
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

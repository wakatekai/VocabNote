VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   6330
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   11310
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim TOTALCOUNT
Dim SlctNm As Integer


Private Sub CommandButton1_Click() 'Aボタン
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 1
    Label5.Caption = Func1(Ans, SlctNm)  '正誤判定結果の画面表示
    ToF = Func1(Ans, SlctNm)             '正誤判定結果取得
End If
End Sub

Private Sub CommandButton2_Click() 'Bボタン
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 2
    Label5.Caption = Func1(Ans, SlctNm)  '正誤判定結果の画面表示
    ToF = Func1(Ans, SlctNm)             '正誤判定結果取得
End If
End Sub

Private Sub CommandButton3_Click() 'Cボタン
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 3
    Label5.Caption = Func1(Ans, SlctNm)  '正誤判定結果の画面表示
    ToF = Func1(Ans, SlctNm)             '正誤判定結果取得
End If
End Sub

Private Sub CommandButton4_Click() 'Dボタン
If TOTALCOUNT = 0 Or TOTALCOUNT = 2 Then
    TOTALCOUNT = 1
    SlctNm = 4
    Label5.Caption = Func1(Ans, SlctNm)  '正誤判定結果の画面表示
    ToF = Func1(Ans, SlctNm)             '正誤判定結果取得
End If
End Sub

Private Sub CommandButton5_Click()  '次の問題へボタン
If TOTALCOUNT = 0 Then
    MsgBox "スキップしますがよろしいですか？"
    TOTALCOUNT = 2
ElseIf TOTALCOUNT = 1 Or TOTALCOUNT = 2 Then
    Unload UserForm4
'    UserForm8.Show
End If
'MsgBox "スキップしますがよろしいですか？"
End Sub


Private Sub CommandButton6_Click()  'タイトル画面へボタン
'If TOTALCOUNT = 3 Then
'    Unload UserForm4
'    UserForm3.Show
'End If
'If TOTALCOUNT = 0 Or 1 Then
'    MsgBox "タイトル画面へ戻りますがよろしいですか？"
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

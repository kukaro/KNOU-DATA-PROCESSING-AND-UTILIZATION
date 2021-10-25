VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 출생연도계산 
   Caption         =   "출생연도계산"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "출생연도계산.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "출생연도계산"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    age = TextBox1.Value
    birthYear = Year(Now) - age
    MsgBox "출생년도는 " & birthYear & "년 입니다."
End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub

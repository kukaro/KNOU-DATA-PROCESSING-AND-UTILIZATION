VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���������� 
   Caption         =   "UserForm1"
   ClientHeight    =   1110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "217p.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    age = TextBox1.Value
    birthyear = Year(Now) - age
    MsgBox "��������� " & birthyear & "�� �Դϴ�."
End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

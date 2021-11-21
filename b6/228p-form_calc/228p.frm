VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 사용자정의폼2 
   Caption         =   "UserForm1"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   OleObjectBlob   =   "228p.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "사용자정의폼2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    TextBox2.Font.Bold = CheckBox1.Value
End Sub

Private Sub CheckBox2_Click()
    TextBox2.Font.Italic = CheckBox2.Value
End Sub

Private Sub CheckBox3_Click()
    TextBox2.Font.Underline = CheckBox3.Value
End Sub

Private Sub ComboBox1_Change()
    myForeColor = Switch(ComboBox1.Value = "색을 선택하시오", vbBlack, _
                        ComboBox1.Value = "Red", vbRed, _
                        ComboBox1.Value = "Green", vbGreen, _
                        ComboBox1.Value = "Yellow", vbYellow, _
                        ComboBox1.Value = "Blue", vbBlue, _
                        ComboBox1.Value = "Cyan", vbCyan, _
                        ComboBox1.Value = "Magenta", vbMagenta)
    TextBox2.ForeColor = myForeColor
End Sub

Private Sub CommandButton1_Click()
    Dim intnn As Integer
    Dim power As Integer
    
    Call CheckBox1_Click
    Call CheckBox2_Click
    Call CheckBox3_Click
    Call ComboBox1_Change
    Call ListBox1_Click
    
    nn = TextBox1.Value
    intnn = CInt(nn)
    
    If (OptionButton1.Value = True) Then power = 2
    If (OptionButton2.Value = True) Then power = 3
    
    thisSum = mySum(intnn, power)
    TextBox2.Value = "1^" & power & "부터 " & nn & "^" & power & _
                    "까지의 합은 " & thisSum & "입니다"
    
End Sub

Private Sub CommandButton2_Click()
    Unload 사용자정의폼2
End Sub

Private Sub ListBox1_Click()
    myFontSize = Choose(ListBox1.ListIndex + 1, 10, 12, 14, 16, 18, 20)
    TextBox2.Font.Size = myFontSize
End Sub

Private Sub OptionButton1_Click()
    Call CommandButton1_Click
End Sub

Private Sub OptionButton2_Click()
    Call CommandButton2_Click
End Sub


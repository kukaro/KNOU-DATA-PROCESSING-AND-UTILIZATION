Attribute VB_Name = "Module1"
Sub myForm2()
    사용자정의폼2.TextBox1.Value = 10
    사용자정의폼2.ComboBox1.List = Array("Red", "Green", "Yellow", _
                                        "Blue", "Cyan", "Magenta")
    사용자정의폼2.ComboBox1.Text = "색을 선택하시오"
    사용자정의폼2.CheckBox2.Value = True
    사용자정의폼2.ListBox1.List = Array("10pt", "12pt", "14pt", _
                                        "16pt", "18pt", "20pt")
    사용자정의폼2.ListBox1.ListIndex = 1
    사용자정의폼2.OptionButton1.Value = True
    사용자정의폼2.Show
End Sub

Function mySum(nn As Integer, power As Integer) As Integer
    mySum = 0
    For ia = 1 To nn
        mySum = mySum + ia ^ power
    Next ia
End Function

Attribute VB_Name = "Module1"
Sub 절대참조제곱()
Attribute 절대참조제곱.VB_Description = "C2부터 C11에 B2~B11의 제곱계산"
Attribute 절대참조제곱.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 절대참조제곱 매크로
' C2부터 C11에 B2~B11의 제곱계산
'
' 바로 가기 키: Ctrl+a
'
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    Selection.AutoFill Destination:=Range("C2:C11"), Type:=xlFillDefault
    Range("C2:C11").Select
End Sub

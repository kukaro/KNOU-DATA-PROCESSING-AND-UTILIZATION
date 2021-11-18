Attribute VB_Name = "Module1"
Sub 상대참조제곱()
Attribute 상대참조제곱.VB_Description = "상대 참조로 왼쪽열의 제곱 계산"
Attribute 상대참조제곱.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' 상대참조제곱 매크로
' 상대 참조로 왼쪽열의 제곱 계산
'
' 바로 가기 키: Ctrl+a
'
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A10"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:A10").Select
End Sub

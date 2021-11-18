Attribute VB_Name = "Module1"
Sub 정렬()
Attribute 정렬.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 정렬 매크로
'

'
    Range("B5:H17").Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("B6:B17") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("B5:H17")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 부분합()
Attribute 부분합.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 부분합 매크로
'

'
    Range("B5:H17").Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(6, 7), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
End Sub
Sub 부분합삭제()
Attribute 부분합삭제.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 부분합삭제 매크로
'

'
    Range("B5:H21").Select
    Selection.RemoveSubtotal
End Sub

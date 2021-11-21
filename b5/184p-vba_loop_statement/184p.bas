Attribute VB_Name = "Module1"
Sub WhileTest1()
    i = 1
    Do While i < 10
        Cells(i, 1) = i * 10
        i = i + 1
    Loop
    MsgBox "데이터 시트에서 변한 값을 확인하세요"
End Sub

Sub WhileTest2()
    i = 10
    Do Until i < 1
        Cells(1, i) = i * 10
        i = i - 1
    Loop
    MsgBox "데이터 시트에서 변한 값을 확인하세요"
End Sub

Sub DoLoopTest2()
    myno = 1
    Do Until myno = 10
        If (Not IsNumeric(Cells(1, myno).Value)) Then
            MsgBox (myno & "번째 값이 숫자가 아님")
            Cells(2, myno) = "중지"
            Exit Do
        Else
            Cells(2, myno) = Cells(1, myno)
        End If
        myno = myno + 1
    Loop
End Sub

Sub ForEachTest()
    For Each myCell In Range("A1:A5")
        myCell.Value = myCell.Value ^ 2
    Next
End Sub

Sub myClear()
    myNumber = Array(1, 2, 3, 4, 5)
    For Each ii In myNumber
        Cells(ii, 1).Value = ii
    Next
End Sub

Sub myInit()
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    For Each i In Array(1, 2, 3, 4, 5)
        ws1.Copy ThisWorkbook.Sheets(1)
    Next
End Sub

Sub ForEachTest2()
    nSheet = ThisWorkbook.Sheets.Count
    For Each ws In ThisWorkbook.Worksheets
        If nSheet = 1 Then Exit For
        ws.Select
        ActiveWindow.SelectedSheets.Delete
        nSheet = nSheet - 1
    Next
End Sub

Sub ForTest1()
    Dim Loc As Integer
    For i = 0 To 255 Step 5
        Loc = (i \ 5) + 1
        Cells(Loc, 1).Interior.Color = RGB(i, i, i)
        Cells(Loc, 2).Interior.Color = RGB(i, 0, 0)
        Cells(Loc, 3).Interior.Color = RGB(0, i, 0)
        Cells(Loc, 4).Interior.Color = RGB(0, 0, i)
    Next i
End Sub

Sub Beeps()
    Application.EnableSound = True
    For ix = 1 To 10
        Application.Wait (Now() + CDate("00:00:01"))
        Beep
    Next ix
End Sub

Sub ForTest3()
Application.EnableSound = False
    For i = 5 To 1 Step -1
        For j = 1 To 5
            Cells(i, j) = i ^ j
            Cells(i, j).Interior.ColorIndex = i + 1
        Next j
    Next i
End Sub

Sub ForTest4()
    MsgBox "" & Cells(1, 6)
    For i = 5 To 1 Step -1
        For j = 1 To 5
            If (j > Cells(1, 6)) Then Exit For
            Cells(i, j) = i ^ j
            Cells(i, j).Interior.ColorIndex = i + 1
        Next j
    Next i
End Sub

Sub WithTest1()
    With Worksheets("Sheet1").Range("A1:F6")
        .Value = 30
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0)
    End With
End Sub

Sub WithTest2()
    With Worksheets("Sheet1").Cells(1, 1)
        .Formula = "=SQRT(50)"
        With .Font
            .Name = "Arial"
            .Bold = True
            .Size = 18
        End With
    End With
End Sub

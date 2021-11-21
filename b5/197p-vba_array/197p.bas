Attribute VB_Name = "Module1"
Sub 배열1()
    Dim myArray, myArray2
    myArray = Array("가", "나", "다")
    Range("A1:C1") = myArray
    '------------------------------
    myArray2 = Split("A|B|C|1|2|3", "|", -1, vbBinaryCompare)
    Range("A2:F2") = myArray2
    '------------------------------
    Dim test(9) As Integer
    For i = 0 To 9
        test(i) = i * 10
        Cells(3, i + 1) = test(i)
    Next i
    '------------------------------
    Dim myMtx(1 To 3, 1 To 3)
    For irow = 1 To 3
        For icol = 1 To 3
            myMtx(irow, icol) = irow * 10 + icol
            Cells(irow + 4, icol) = myMtx(irow, icol)
        Next icol
    Next irow
    '------------------------------
    Dim myArray3() As Integer
    ReDim myArray3(1 To 10)
    myArray3 = test
    For i = 0 To 9
        Cells(9, i + 1) = myArray3(i)
    Next i
End Sub

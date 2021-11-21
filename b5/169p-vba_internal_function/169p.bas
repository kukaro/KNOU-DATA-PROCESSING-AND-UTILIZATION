Attribute VB_Name = "Module1"
Sub 숫자()
    For ia = 1 To 10
        Cells(1, ia) = Rnd()
    Next ia
    
    Cells(2, 1) = Abs(-23.45)
    ' 소수점 이하 그냥 버림
    Cells(3, 1) = Fix(-23.45)
    ' 해당 값보다 크지않은 정수중 가장 큰 정수
    Cells(4, 1) = Int(-23.45)
    Cells(5, 1) = Round(-23.45)
End Sub

Sub 문자열()
    For ia = 1 To 10
        Cells(1, ia) = Chr(ia + 55)
    Next ia
    
    aa = " 내 문자열이다"
    bb = "My String Text "
    Cells(2, 1) = Len(aa): Cells(2, 2) = Len(bb)
    Cells(3, 1) = LCase(bb): Cells(3, 2) = UCase(bb)
    Cells(4, 1) = String(5, "Y")
    Cells(5, 1) = Val("-23.45")
    Cells(6, 1) = LTrim(aa)
    Cells(7, 1) = " A": Cells(7, 2) = Trim(" A")
End Sub

Sub 날짜시간1()
    Dim myLabels
    myLabels = Array("Now", "연도", "월", "일", "요일", _
        "오늘날짜", "시", "분", "초", "현재시각")
    With Range("A1:J1")
        .Clear
        .Cells.Interior.Color = RGB(0, 255, 255)
        .Value = myLabels
        .Font.Bold = True
    End With
    
    현재시간 = Now
    Cells(2, 1) = 현재시간
    y = Year(현재시간): Cells(2, 2) = y
    m = Month(현재시간): Cells(2, 3) = m
    d = Day(현재시간): Cells(2, 4) = d
    Cells(2, 5) = Choose(Weekday(현재시간), _
        "일", "월", "화", "수", "목", "금", "토")
    오늘날짜 = DateSerial(y, m, d): Cells(2, 6) = 오늘날짜
    h = Hour(현재시간): Cells(2, 7) = h
    m = Minute(현재시간): Cells(2, 8) = m
    s = Second(현재시간): Cells(2, 9) = s
    현재시간 = TimeSerial(h, m, s): Cells(2, 10) = 현재시간
End Sub

Sub 개체확인()
    Dim myLabels
    myLabels = Array("Date", "연도", "월", "일", "분기")
    현재시각 = Now
    ia = 1
    xx = 1.2345
    
    Cells(1, 1) = IsArray(myLabels)
    Cells(1, 2) = IsArray(ia)
    Cells(1, 3) = IsEmpty(ia)
    Cells(1, 4) = IsNull(ia)
    Cells(1, 5) = IsDate("2020-1-1")
    Cells(1, 6) = IsObject(ia)
    Cells(1, 7) = IsNull(현재시각)
End Sub

Sub fixdate0()
    Cells.Clear
    mydate = #2/13/2020#
    If mydate < Now Then mydate = Now
    Cells(1, 1) = mydate
End Sub

Sub fixdate()
    Cells.Clear
    mydate = #1/1/2020#
    If mydate < Now Then
        mydate = Now
        Cells(1, 1) = mydate
    End If
End Sub

Sub IfElseTest1()
    If (Range("A1").Value > 0) Then
        Range("B1").Value = "양수"
        Range("B1").Interior.Color = RGB(0, 255, 255)
    Else
        Range("B1").Value = "음수"
        Range("B1").Interior.Color = RGB(255, 255, 0)
    End If
End Sub

Sub IfElseTest2()
    If (Range("A1").Value > 0) Then
        Range("B1").Value = "양수"
        Range("B1").Interior.Color = RGB(0, 255, 255)
    ElseIf (Range("A1").Value = 0) Then
        Range("B1").Value = "Zero"
        Range("B1").Interior.Color = RGB(255, 0, 255)
    Else
        Range("B1").Value = "음수"
        Range("B1").Interior.Color = RGB(255, 255, 0)
    End If
End Sub

Sub Iiftest()
    For i = 1 To 10
        Cells(i, 2) = IIf(Cells(i, 1).Value > 0, "양수", "그 외")
    Next i
End Sub

Sub SelectCaseTest1()
    remainder = Range("A1").Value Mod 3
    Range("B1").Cells.Font.Color = vbRed
    Select Case remainder
        Case 0
            Range("B1").Value = " 3의 배수"
        Case 1
            Range("B1").Value = " 3의 배수+1"
        Case 2
            Range("B1").Value = " 3의 배수+2"
    End Select
End Sub

Sub SwitchTest()
    공급자번호 = Range("A1").Value
    Dim 공급자이름 As String
    공급자이름 = Switch(공급자번호 = 1, "IBM", _
                공급자번호 = 2, "HP", _
                공급자번호 = 3, "NVIDIA")
    Range("A2").Value = 공급자이름
End Sub

Attribute VB_Name = "Module1"
Sub 입력상자()
    num = InputBox("아무거나 입력하세요", "입력상자보기", "여기에 입력", 1, 1)
    Range("A1").Value = num
End Sub

Sub 메시지상자보기()
    msg = "계속 진행할까요?"
    Title = "메시지 상자 보기"
    Style = vbInformation + vbYesNoCancel + vbDefaultButton2
    response = MsgBox(msg, Style, Title)
    If response = vbYes Then
        response2 = MsgBox("예를 클릭하였습니다.", vbOKOnly, Title)
    Else
        response2 = MsgBox("아니요 또는 취소를 클릭하였습니다.", vbOKOnly, Title)
    End If
End Sub

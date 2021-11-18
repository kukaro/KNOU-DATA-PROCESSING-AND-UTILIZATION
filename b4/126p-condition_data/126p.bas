Attribute VB_Name = "Module1"
Sub 중부영업소()
Attribute 중부영업소.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 중부영업소 매크로
'

'
    Rows("8:75").Select
    Selection.EntireRow.Hidden = False
    Rows("32:75").Select
    Selection.EntireRow.Hidden = True
End Sub

Sub 서부영업소()
'
' 서부영업소 매크로
'

'
    Rows("8:75").Select
    Selection.EntireRow.Hidden = False
    Rows("8:31").Select
    Selection.EntireRow.Hidden = True
    Rows("51:75").Select
    Selection.EntireRow.Hidden = True
End Sub

Sub 동부영업소()
'
' 동부영업소 매크로
'

'
    Rows("8:75").Select
    Selection.EntireRow.Hidden = False
    Rows("8:50").Select
    Selection.EntireRow.Hidden = True
    Rows("75:75").Select
    Selection.EntireRow.Hidden = True
End Sub

Sub 전체보기()
'
' 전체보기 매크로
'

'
    Rows("8:75").Select
    Selection.EntireRow.Hidden = False
End Sub



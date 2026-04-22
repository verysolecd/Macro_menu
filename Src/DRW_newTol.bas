Attribute VB_Name = "DRW_newTol"
Private Const mdlname As String = "DRW_newTol"
Sub newTol()


Set oDrw = CATIA.ActiveDocument
Set rtDrw = oDrw.DrawingRoot
Set shts = rtDrw.sheets
Set osht = shts.item(1)
Set oVs = osht.Views
Set oView = oVs.ActiveView

Set ogdt = oView.GDTs.item(1) 'Add(1, 1, 20, 20, 10, "00")
tex = ogdt.GetReferenceNumber(1)
Set ogdt = oView.GDTs.Add(1, 1, 20, 20, 10, "1ABC")
Set colls = ogdt.Leaders

On Error Resume Next
  Call colls.Remove(1)
On Error GoTo 0
Call ogdt.SetToleranceType(1, 10)
Dim istart

Dim iend

Set txt = ogdt.GetTextRange(1, 1)


'1  直线度
'2  平面度
'3  圆度
'4  圆柱度
'5  线轮廓度
'6  面轮廓度
'7  角度
'8  垂直度
'9  平行度
'10  位置度
'11  同轴度
'12  对称度
'13  圆跳动
'14  全跳动
'15

End Sub


Dim ws As Worksheet
Dim i&, rng As Range
Dimp As Shape图片是Shape对象中的一种Dimpas，pns”图片路径和名字=====当前工作表
Set ws = ActiveSheet
With ws.
清除之前的内容和形状. Range ("B:B"). Clear
For Each p In . Shapes
p. Delete Next p
For i = 2 To . Cells(. Rows. Count, 1).End (3). Row pn = . Cells(i, 1). Value
pa = ThisWorkbook. Path & "\" & pn & ". jpg"
If Dir(pa) = "" Then
.Cells(i，2). Value ="图片不存在"
Else
先插入图片：路径，不链接，保存图片，-1保留原有高度宽度Set p = . Shapes.AddPicture (pa, False, True, O, O, -1, -1)Set rng = . Cells(i, 2)
==保持纵横比完整自适应DS写的=
With p
保持纵横比
. LockAspectRatio = msoTrue计算最佳尺寸以适应单元格
Dim cellRatio As Double, imgRatio As Double cellRatio = rng. Width / rng. Height imgRatio = . Width / . Height
If imgRatio > cellRatio Then
图片较宽，以单元格宽度为准.Width = rng. Width
.Top = rng.Top+（rng.Height－.Height) V 2’垂直居中. Left = rng. Left
图片较高，以单元格高度为准.Height = rng. Height. Top = rng. Top
.Left = rng.Left +（rng. Width -.Width) / 2 ′ 水平居中End If
. Placement = xlMoveAndSize
End With
End If
Next i
End With
MsgBox "ok"
End Sub
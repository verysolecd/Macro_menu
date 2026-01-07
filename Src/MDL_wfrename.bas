Attribute VB_Name = "MDL_wfrename"
'Attribute VB_Name = "m24_wfrename"
' 重命名几何元素
'{GP:4}
'{EP:wfrename}
'{Caption:几何重命名}
'{ControlTipText: 提示选择几何图形集后导出下面的点集}
'{BackColor:12648447}
' = 0 , Unknown
' = 1 , Point
' = 2 , Curve
' = 3 , Line
' = 4 , Circle
' = 5 , Surface
' = 6 , Plane
' = 7 , Solid, Volume

Sub wfrename()
   
  If CATIA.Windows.count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
    
    Dim oDoc
    On Error Resume Next
        Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0
    Dim str
    str = TypeName(oDoc)
    If Not str = "PartDocument" Then
    MsgBox "没有打开的part"
    Exit Sub
    End If
 

End Sub

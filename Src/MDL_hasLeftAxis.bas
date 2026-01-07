Attribute VB_Name = "MDL_hasLeftAxis"
'Attribute VB_Name = "MDL_hasLeftAxis"

' 检查零件文档中是否存在左手坐标系
'{Gp:999}
'{Ep:LeftHand}
'{Caption:LeftHandAxis}
'{ControlTipText:检查是否有左手坐标系}
'{BackColor:33023}
Option Explicit
Sub LeftHand()
    ' 检查是否可以执行
    If Not CanExecute("PartDocument") Then Exit Sub
    Dim doc As PartDocument: Set doc = CATIA.ActiveDocument
    Dim Axs As AxisSystems: Set Axs = doc.part.AxisSystems
    Dim ax As AxisSystem
    Dim msg As String: msg = vbNullString
    For Each ax In Axs
        If IsLeft(ax) Then
            msg = msg & ax.Name & vbNewLine
        End If
    Next
    If msg = vbNullString Then
        MsgBox "未找到左手坐标系。"
    Else
        MsgBox "已找到左手坐标系：" & vbNewLine & msg
    End If
End Sub

Private Function IsLeft(ByVal ax As Variant) As Boolean
    ' 定义向量
    Dim vecX(2), vecY(2), VecZ(2)
    ax.GetXAxis vecX
    ax.GetYAxis vecY
    ax.GetZAxis VecZ
    
    ' 计算 X 轴和 Y 轴的叉积
    Dim Outer(2) As Double
    Outer(0) = vecX(1) * vecY(2) - vecX(2) * vecY(1)
    Outer(1) = vecX(2) * vecY(0) - vecX(0) * vecY(2)
    Outer(2) = vecX(0) * vecY(1) - vecX(1) * vecY(0)
    
    ' 计算叉积结果与 Z 轴的点积，并判断是否小于 0
    IsLeft = _
        VecZ(0) * Outer(0) + VecZ(1) * Outer(1) + VecZ(2) * Outer(2) < 0
End Function


Attribute VB_Name = "m_hasLeftAxis"
'Attribute VB_Name = "sample_IsExists_LeftAxis"

' 检查零件文档中是否存在左手坐标系
'{Gp:1}
'{Ep:LeftHand}
'{Caption:LeftHandAxis}
'{ControlTipText:检查是否有左手坐标系}
'{BackColor:33023}
Option Explicit
Sub LeftHand()
    ' 检查是否可以执行
    If Not CanExecute("PartDocument") Then Exit Sub
    Dim Doc As PartDocument: Set Doc = CATIA.ActiveDocument
    Dim Axs As AxisSystems: Set Axs = Doc.Part.AxisSystems
    Dim Ax As AxisSystem
    Dim Msg As String: Msg = vbNullString
    For Each Ax In Axs
        If IsLeft(Ax) Then
            Msg = Msg & Ax.Name & vbNewLine
        End If
    Next
    If Msg = vbNullString Then
        MsgBox "未找到左手坐标系。"
    Else
        MsgBox "已找到左手坐标系：" & vbNewLine & Msg
    End If
End Sub

Private Function IsLeft(ByVal Ax As Variant) As Boolean
    ' 定义向量
    Dim VecX(2), VecY(2), VecZ(2)
    Ax.GetXAxis VecX
    Ax.GetYAxis VecY
    Ax.GetZAxis VecZ
    
    ' 计算 X 轴和 Y 轴的叉积
    Dim Outer(2) As Double
    Outer(0) = VecX(1) * VecY(2) - VecX(2) * VecY(1)
    Outer(1) = VecX(2) * VecY(0) - VecX(0) * VecY(2)
    Outer(2) = VecX(0) * VecY(1) - VecX(1) * VecY(0)
    
    ' 计算叉积结果与 Z 轴的点积，并判断是否小于 0
    IsLeft = _
        VecZ(0) * Outer(0) + VecZ(1) * Outer(1) + VecZ(2) * Outer(2) < 0
End Function

 BaseScene As Variant: BaseScene = GetScene3D(GetViewPnt3D())
    Dim TopDoc As ProductDocument: Set TopDoc = CATIA.Documents.Add("Product")
    Call ToProdu
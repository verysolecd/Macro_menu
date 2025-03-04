Attribute VB_Name = "m2_hasLeftAxis"
'Attribute VB_Name = "m2_hasLeftAxis"

' �������ĵ����Ƿ������������ϵ
'{Gp:2}
'{Ep:LeftHand}
'{Caption:LeftHandAxis}
'{ControlTipText:����Ƿ�����������ϵ}
'{BackColor:33023}
Option Explicit
Sub LeftHand()
    ' ����Ƿ����ִ��
    If Not CanExecute("PartDocument") Then Exit Sub
    Dim Doc As PartDocument: Set Doc = CATIA.ActiveDocument
    Dim Axs As AxisSystems: Set Axs = Doc.part.AxisSystems
    Dim Ax As AxisSystem
    Dim Msg As String: Msg = vbNullString
    For Each Ax In Axs
        If IsLeft(Ax) Then
            Msg = Msg & Ax.Name & vbNewLine
        End If
    Next
    If Msg = vbNullString Then
        MsgBox "δ�ҵ���������ϵ��"
    Else
        MsgBox "���ҵ���������ϵ��" & vbNewLine & Msg
    End If
End Sub

Private Function IsLeft(ByVal Ax As Variant) As Boolean
    ' ��������
    Dim VecX(2), VecY(2), VecZ(2)
    Ax.GetXAxis VecX
    Ax.GetYAxis VecY
    Ax.GetZAxis VecZ
    
    ' ���� X ��� Y ��Ĳ��
    Dim Outer(2) As Double
    Outer(0) = VecX(1) * VecY(2) - VecX(2) * VecY(1)
    Outer(1) = VecX(2) * VecY(0) - VecX(0) * VecY(2)
    Outer(2) = VecX(0) * VecY(1) - VecX(1) * VecY(0)
    
    ' ����������� Z ��ĵ�������ж��Ƿ�С�� 0
    IsLeft = _
        VecZ(0) * Outer(0) + VecZ(1) * Outer(1) + VecZ(2) * Outer(2) < 0
End Function


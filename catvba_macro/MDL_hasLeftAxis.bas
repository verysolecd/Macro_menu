Attribute VB_Name = "MDL_hasLeftAxis"
'Attribute VB_Name = "m2_hasLeftAxis"

' �������ĵ����Ƿ������������ϵ
'{Gp:999}
'{Ep:LeftHand}
'{Caption:LeftHandAxis}
'{ControlTipText:����Ƿ�����������ϵ}
'{BackColor:33023}
Option Explicit
Sub LeftHand()
    ' ����Ƿ����ִ��
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
        MsgBox "δ�ҵ���������ϵ��"
    Else
        MsgBox "���ҵ���������ϵ��" & vbNewLine & msg
    End If
End Sub

Private Function IsLeft(ByVal ax As Variant) As Boolean
    ' ��������
    Dim vecX(2), vecY(2), VecZ(2)
    ax.GetXAxis vecX
    ax.GetYAxis vecY
    ax.GetZAxis VecZ
    
    ' ���� X ��� Y ��Ĳ��
    Dim Outer(2) As Double
    Outer(0) = vecX(1) * vecY(2) - vecX(2) * vecY(1)
    Outer(1) = vecX(2) * vecY(0) - vecX(0) * vecY(2)
    Outer(2) = vecX(0) * vecY(1) - vecX(1) * vecY(0)
    
    ' ����������� Z ��ĵ�������ж��Ƿ�С�� 0
    IsLeft = _
        VecZ(0) * Outer(0) + VecZ(1) * Outer(1) + VecZ(2) * Outer(2) < 0
End Function


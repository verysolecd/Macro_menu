Attribute VB_Name = "ASM_copychildren"
'Attribute VB_Name = "M36_copychildren"
' ����
'{GP:3}
'{EP:cpChildren}
'{Caption:�����Ӳ�Ʒ}
'{ControlTipText: һ�����Ƶ�һ����Ʒ���Ӳ�Ʒ���ڶ�����Ʒ�Ӽ�}
'{BackColor:}
' ����ģ�鼶����

Sub cpChildren()
If CATIA.Windows.Count < 1 Then
    MsgBox "û�д򿪵Ĵ���"
    Exit Sub
End If
If Not CanExecute("ProductDocument") Then Exit Sub

Dim imsg, filter(0), iSel
Set oDoc = CATIA.ActiveDocument
Set osel = CATIA.ActiveDocument.Selection

On Error Resume Next
    imsg = "��ѡ��Ҫ���Ƶ��Ӳ�Ʒ����"
    MsgBox imsg
    filter(0) = "Product"
    Dim sourcePrd, targetPrd
    Set sourcePrd = KCL.SelectElement(imsg, filter).Value
    If sourcePrd Is Nothing Then Exit Sub
    For Each Prd In sourcePrd.Products
       osel.Add Prd
    Next
        osel.Copy
        osel.Clear
    imsg = "��ѡ�����Ŀ���Ʒ"
    MsgBox imsg
    Set targetPrd = KCL.SelectElement(imsg, filter).Value
    If targetPrd Is Nothing Then
        Exit Sub
    Else
        osel.Add targetPrd
        osel.Paste
        Set targetPrd = Nothing
    End If
On Error GoTo 0
End Sub


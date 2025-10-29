Attribute VB_Name = "ASM_deleteChildren"
'Attribute VB_Name = "M37_DeleteChildren"
' ����
'{GP:3}
'{EP:DeleteChildren}
'{Caption:ɾ���Ӳ�Ʒ}
'{ControlTipText: һ��ɾ��ѡ��Ĳ�Ʒ���Ӳ�Ʒ}
'{BackColor:}
' ����ģ�鼶����
Option Explicit

Sub DeleteChildren()
  If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Sub
  End If
  If Not CanExecute("ProductDocument") Then Exit Sub
    Dim osel: Set osel = CATIA.ActiveDocument.Selection: osel.Clear

    Dim imsg, filter(0), iSel
      imsg = "��ѡ�񸸼�": filter(0) = "Product"
       Set iSel = KCL.SelectItem(imsg, filter)
    If iSel Is Nothing Then Exit Sub
    Dim prd
    For Each prd In iSel.Products
      osel.Add prd
    Next
      Dim btn, bTitle, bResult
      imsg = "��ɾ��" & iSel.PartNumber & iSel.Name & "�������Ӳ�Ʒ����ȷ����"
      btn = vbYesNo + vbExclamation
      bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
           Select Case bResult
              Case 7: Exit Sub '===ѡ�񡰷�====
              Case 6  '===ѡ���ǡ�,���в�Ʒѡ��====
                  On Error Resume Next
                       osel.Delete
                       osel.Clear
                  On Error GoTo 0
          End Select

End Sub

Attribute VB_Name = "ASM_ChildMng"
'------宏信息---------------------------------
'{GP:3}
'{EP:ChildMng}
'{Caption:子产品管理}
'{ControlTipText: 复制黏贴和删除子产品}
'{BackColor:}
'------控件信息------------------------------
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI Label lbL_tip 点击功能后,依次选择源产品→目标产品
' %UI Button btn_copy 复制后黏贴
' %UI Button btn_delete 删除子件
' %UI Button btn_cancel 取消
'------其他------------------------------

Private Const mdlname As String = "ASM_ChildMng"
Sub ChildMng()
    If Not CanExecute("ProductDocument") Then Exit Sub
    Dim oFrm: Set oFrm = KCL.newFrm("ASM_ChildMng")
              Select Case oFrm.BtnClicked
                Case "btn_copy"
                    Call cpChildren
                Case "btn_delete"
                    Call DeleteChildren
                 Case Else
                    Exit Sub
            End Select
End Sub


Sub cpChildren()
    Dim imsg, filter(0), oSel
    Set oDoc = CATIA.ActiveDocument
    Set oSel = CATIA.ActiveDocument.Selection
    oSel.Clear
    On Error GoTo ErrorHandler
        Call KCL.setASM(False)
        imsg = "请先点击选择源父产品，再点击选择目标父产品": MsgBox imsg
        filter(0) = "Product"
        Dim sourcePrd, targetPrd As Product
        Set sourcePrd = KCL.SelectItem(imsg, filter)
        If sourcePrd Is Nothing Then GoTo ErrorHandler
            For Each prd In sourcePrd.Products
               oSel.Add prd
            Next
        oSel.Copy: oSel.Clear
        imsg = "请点击选择目标父产品"
        Set targetPrd = KCL.SelectItem(imsg, filter)
        If targetPrd Is Nothing Then
          GoTo ErrorHandler
        Else
            oSel.Add targetPrd
            oSel.Paste
        End If
            oSel.Clear
            Set targetPrd = Nothing
            Set sourcePrd = Nothing
            Call KCL.setASM(True)
    On Error GoTo 0

ErrorHandler:
        If Err.Number <> 0 Then
            Call KCL.setASM(True)
              oSel.Clear
            MsgBox "CATIA 程序错误：" & Err.Description, vbCritical
                 Err.Clear
            Exit Sub
        Else
         Call KCL.setASM(True)
        End If
         Call KCL.setASM(True)
End Sub
Sub DeleteChildren()
    Dim oSel: Set oSel = CATIA.ActiveDocument.Selection: oSel.Clear
    Dim imsg, filter(0), iSel
      imsg = "请选择父集": filter(0) = "Product"
       Set iSel = KCL.SelectItem(imsg, filter)
    If iSel Is Nothing Then Exit Sub
    Dim prd
    For Each prd In iSel.Products
      oSel.Add prd
    Next
      Dim btn, bTitle, bResult
      imsg = "将删除" & iSel.partNumber & iSel.Name & "下的所有子产品，您确认吗"
      btn = vbYesNo + vbExclamation
      bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
           Select Case bResult
              Case 7: Exit Sub '===选择“否”====
              Case 6  '===选择“是”,进行产品选择====
                  On Error Resume Next
                       oSel.Delete
                       oSel.Clear
                  On Error GoTo 0
          End Select
End Sub

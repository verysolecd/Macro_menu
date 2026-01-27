Attribute VB_Name = "ASM_2Localsend"
'{GP:3}
'{Ep:sendDir}
'{Caption:备份到路径}
'{ControlTipText:send当前根产品到路径}
'{BackColor:}

Private Const mdlname As String = "ASM_localsend"
Sub sendDir()
    If Not CanExecute("ProductDocument,DrawingDocument") Then Exit Sub
    CATIA.DisplayFileAlerts = True
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    ipath_name = oDoc.path & "\" & oDoc.Name
    Dim opath
        opath = KCL.ofParentPath(oDoc.path)
    Dim pn
        If KCL.IsObj_T(oDoc, "DrawingDocument") Then
            pn = KCL.strbflast(oDoc.Name, ".")
        Else
            pn = oDoc.Product.partNumber
        End If
    Dim bckFolderName As String
    fname = KCL.rmchn(pn)    '将零件号所有中文字符替换为" "
    bckFolderName = KCL.strbflast(fname, "_") & "_" & KCL.timestamp("min")
    bckpath = opath & bckFolderName
    
    If KCL.isExists(oDoc.path) Then
        Dim btn, bTitle, bResult
            imsg = "将备份到" & bckpath & "您确认吗？"
            btn = vbYesNo + vbExclamation
            bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
            Select Case bResult
                Case 7: Exit Sub '===选择“否”====
                Case 2: Exit Sub '===选择“取消”====
                Case 6  '===选择“是”====
                    If Not KCL.isPathchn(bckpath) Then
                        Set Send = CATIA.CreateSendTo()
                        Send.KeepDirectory (1)  '1 keepdir ， 0 no keep dir
                        Send.SetInitialFile ipath_name
                        Send.SetDirectoryFile bckpath
                        Send.Run
                        MsgBox "已经备份到" & bckpath
                  Else
                      MsgBox bckFolder & vbNewLine & _
                      "  " & vbNewLine & _
                      "你的产品零件号包含非法字符，无法备份，请检查!"
                  End If
            End Select
    Else
        MsgBox bckFolder & vbNewLine & _
        "  " & vbNewLine & _
        "你的产品路径不存在，无法备份，请检查!"
    End If
End Sub


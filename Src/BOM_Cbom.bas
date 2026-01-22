Attribute VB_Name = "BOM_Cbom"
'------宏信息-----------------------------------------------------
'{GP:2}
'{Ep:cBom}
'{Caption:生成BOM}
'{ControlTipText:一键生成带有截图的BOM}
'{BackColor:16744703}
'------弹窗控件----------------------------------------------------
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_capture   是否同时截图到catia
' %UI CheckBox chk_GXfmt   是否GX格式
' %UI Button btnOK  生成BOM
' %UI Button btncancel  取消

Sub cBom()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    CATIA.StartCommand ("* iso")
    Set oFrm = KCL.newFrm("BOM_Cbom")
    Select Case oFrm.BtnClicked
        Case "btnOK":
                If pdm Is Nothing Then Set pdm = New Cls_PDM
                If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = pdm.getiPrd()
                Dim iprd: Set iprd = pdm.CurrentProduct
                If iprd Is Nothing Then Exit Sub
                If gws Is Nothing Then Set xlm = New Cls_XLM
                Call Cal_Mass2
                g_counter = 1: lv = 1
                
                If Not oFrm.Res("chk_GXfmt") Then
                    Dim tmpData(): tmpData() = pdm.recurInfoPrd(iprd, lv)
                    
                    ReDim resultAry(1 To UBound(tmpData, 1), 1 To UBound(tmpData, 2) + 2)
                        For i = 1 To UBound(tmpData, 1)
                             For j = 1 To UBound(resultAry, 2)
                               Select Case j
                                    Case 1: resultAry(i, j) = i
                                    Case Else: resultAry(i, j) = tmpData(i, (j - 2))
                               End Select
                             Next j
                     Next i
                     
                     
                    Dim idx, idcol
                        idcol = Array(0, 1, 2, 3, 4, 5, 7, 8, 10, 11, 13) ' 目标列号, 0号元素不占位置
                          idx = Array(0, 1, 2, 3, 4, 5, 11, 9, 7, 10, 7)  ' 需提取属性索引（0-based)
                        xlm.inject_bom resultAry, idcol, idx
                    If oFrm.Res("chk_capture") Then
                          Call Capme
                            Dim Colpn, colPic: Colpn = 3: colPic = 6
                            Call xlm.inject_pic(gPic_Path, Colpn, colPic)
                            GoTo Cleanup
                    End If
                        
                Else
                    xlm.inject_gxbom pdm.gxBom(iprd, 1)
                    Set iprd = Nothing
                    xlm.xlshow
                    xlm.freesheet
                End If
                    
             
              GoTo Cleanup
              
        Case Else: Exit Sub
    End Select
  
  Set oFrm = Nothing
    GoTo Cleanup

Cleanup:
On Error Resume Next
 Set oFrm = Nothing
    xlm.xlshow
   Set iprd = Nothing
   KCL.ClearDir (gPic_Path)
   gPic_Path = ""
      xlm.freesheet
      Error.Clear
      On Error GoTo 0
End Sub








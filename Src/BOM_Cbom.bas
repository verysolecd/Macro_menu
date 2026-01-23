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
Option Explicit
Private Const mdlname As String = "BOM_Cbom"
Sub cBom()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    CATIA.StartCommand ("* iso")
    Dim oFrm: Set oFrm = KCL.newFrm(mdlname)
    If oFrm.BtnClicked <> "btnOK" Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = pdm.getiPrd()
    Dim iprd: Set iprd = pdm.CurrentProduct: If iprd Is Nothing Then Exit Sub
    If gws Is Nothing Then Set xlm = New Cls_XLM
    Call Cal_Mass2
    Dim Lv, i, j, Colpn, colPic, idcol, idrow
    g_counter = 1: Lv = 1
    Dim tmpData(): tmpData() = pdm.recurInfoPrd(iprd, Lv)
    
If Not oFrm.Res("chk_GXfmt") Then
        ReDim resultAry(1 To UBound(tmpData, 1), 1 To UBound(tmpData, 2) + 2)
        For i = 1 To UBound(tmpData, 1)
                 For j = 1 To UBound(resultAry, 2)
                   Select Case j
                        Case 1: resultAry(i, j) = i
                        Case Else: resultAry(i, j) = tmpData(i, (j - 2))
                   End Select
                 Next j
         Next i
            idcol = Array(0, 1, 2, 3, 4, 5, 7, 8, 10, 11, 13) ' 目标列号, 0号元素不占位置
            idrow = Array(0, 1, 2, 3, 4, 5, 11, 9, 7, 10, 7)  ' 需提取属性索引（0-based)
            startrow = 2: Colpn = 3: colPic = 6
            xlm.inject_bom resultAry, idcol, idrow
    Else
            '----这一步做了什么？为bom增加了序号，增加了额外的列，单纯属性+LV+count 为9列，最终增加为11列
              ' 为什么要这样写？因为后面写入excel是按列写入的，所以必须为LV这里制造更多的列
         ReDim resultAry(1 To UBound(tmpData, 1), 1 To UBound(tmpData, 2) + 5)
                    For i = 1 To UBound(tmpData, 1)
                        For j = 1 To UBound(resultAry, 2)
                            Select Case j
                                Case 1: resultAry(i, j) = i
                                Case 2, 3, 4, 5
                                    resultAry(i, j) = ""
                                    If tmpData(i, 0) = j Then resultAry(i, j) = tmpData(i, 0)
                                Case Else: resultAry(i, j) = tmpData(i, (j - 5))  '第6列开始对应的原数组的1~9
                            End Select
                        Next j
                    Next i
          idcol = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
          idrow = Array(0, 1, 2, 3, 4, 5, 6, 0, 8, 7, 0, 14, 0, 12, 0, 10)
          startrow = 5: Colpn = 6: colPic = 8
          xlm.inject_gxbom resultAry, idcol, idrow
            
End If
    If oFrm.Res("chk_capture") Then
      Call Capme
      Call xlm.inject_pic(startrow, Colpn, colPic, gPic_Path)
    End If
       GoTo Cleanup
Cleanup:
On Error Resume Next
    Unload oFrm
     Set oFrm = Nothing
    xlm.xlshow
   Set iprd = Nothing
   KCL.ClearDir (gPic_Path)
   gPic_Path = ""
      xlm.freesheet
      Error.Clear
      On Error GoTo 0
End Sub




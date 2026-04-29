Attribute VB_Name = "RW_Cbom"
'------宏信息-----------------------------------------------------
'{GP:1}
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
Private Const mdlname As String = "RW_Cbom"
Sub cBom()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    CATIA.StartCommand ("* iso")
    Dim oEng: Set oEng = KCL.newEngine(mdlname): oEng.Show
    If oEng.ClickedButton <> "btnOK" Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If IsNothing(pdm.CurrentProduct) Then Set pdm.CurrentProduct = KCL.defPrd
    Dim iprd: Set iprd = pdm.CurrentProduct: If IsNothing(iprd) Then Exit Sub
    If IsNothing(gws) Then Set xlm = New Cls_XLM
    Call Cal_Mass
    Dim i, j, startrow, Colpn, colPic
        Dim bomlns() As Bomline: bomlns = pdm.ProduceBOM(iprd)
    If Not oEng.Results("chk_GXfmt") Then
        startrow = 2: Colpn = 3: colPic = 6
        xlm.inject_Bom ConvertBOM_Standard(bomlns), startrow
        startrow = 2: Colpn = 3: colPic = 6
    Else
        startrow = 5: Colpn = 6: colPic = 8
        xlm.inject_GXbom ConvertBOM_GX(bomlns), startrow
        startrow = 5: Colpn = 6: colPic = 8
    End If
    If oEng.Results("chk_capture") Then
      Call CapPrd(iprd)
      Call xlm.inject_pic(startrow, Colpn, colPic, g_Picpath)
    End If
       GoTo Cleanup
Cleanup:
On Error Resume Next
    Unload oEng: Set oEng = Nothing
    Set iprd = Nothing
    xlm.xlshow
   KCL.ClearDir (g_Picpath)
   g_Picpath = ""
   Error.Clear
   On Error GoTo 0
End Sub
Private Function ConvertBOM_Standard(data() As Bomline) As Variant
 'Dim arr2D As Variant: arr2D = ConvertBOM_Standard(data)
    Dim rowCount As Long: rowCount = UBound(data)
    Dim colCount As Long: colCount = 17
    Dim arr2D As Variant: ReDim arr2D(1 To rowCount, 1 To colCount)
    Dim i As Long
    For i = 1 To rowCount
        With data(i)
            arr2D(i, 1) = i                     ' No. 编号
            arr2D(i, 2) = .level                ' Layout 层级
            arr2D(i, 3) = .partNumber           ' PN 零件号
            arr2D(i, 4) = .Nomenclature         ' Nomenclature 英文名称
            arr2D(i, 5) = .Definition           ' Definition 中文名称
            ' arr2D(i, 6) 图像列，后续填充
            arr2D(i, 7) = .Quantity             ' Quantity 数量
            arr2D(i, 8) = .Mass                 ' Weight 单质量
            ' arr2D(i, 9) 总质量由公式计算
            arr2D(i, 10) = .Material            ' Material 材料
            arr2D(i, 11) = .Thickness           ' Thickness 厚度
            ' arr2D(i, 12) 空列
            arr2D(i, 13) = .Material            ' Material 材料(重复)
            arr2D(i, 14) = .Density             ' Density 密度
            ' arr2D(i, 15-17) 预留给材料属性(抗拉、屈服、延伸率)
        End With
    Next i
    ConvertBOM_Standard = arr2D
End Function
' 转换函数：Bomline() → 二维数组 (GX格式)
Private Function ConvertBOM_GX(data() As Bomline) As Variant
 '   Dim arr2D As Variant: arr2D = ConvertBOM_GX(data)
    Dim rowCount As Long: rowCount = UBound(data)
    Dim colCount As Long: colCount = 16
    Dim arr2D As Variant: ReDim arr2D(1 To rowCount, 1 To colCount)
    Dim i As Long
    For i = 1 To rowCount
        With data(i)
            arr2D(i, 1) = i                                      ' NO.
            ' Level spreading (Cols 2-5) - TIRE列
            arr2D(i, 2) = IIf(.level = 1, 1, "")
            arr2D(i, 3) = IIf(.level = 2, 2, "")
            arr2D(i, 4) = IIf(.level = 3, 3, "")
            arr2D(i, 5) = IIf(.level = 4, 4, "")
            arr2D(i, 6) = .partNumber                            ' PART NO.
            arr2D(i, 7) = ""                                     ' DRAWING NO.
            arr2D(i, 8) = .Definition                            ' 中文名称
            arr2D(i, 9) = .Nomenclature                          ' ITEM NAME
            arr2D(i, 10) = ""                                    ' Ver.
            arr2D(i, 11) = .Quantity                             ' QTY.
            arr2D(i, 12) = ""                                    ' UNIT
            arr2D(i, 13) = .Mass                                 ' UNIT WEIGHT
            arr2D(i, 14) = ""                                       ' TOTAL WEIGHT
            arr2D(i, 15) = .Material                             ' MATERIAL
            arr2D(i, 16) = ""                                    ' 备用列
        End With
    Next i
    ConvertBOM_GX = arr2D
End Function


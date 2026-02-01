Attribute VB_Name = "MDL_MaterialColors"

'------UI Definition-----------------------
'{GP:3}
'{EP:MaterialPainter}
'{Caption:Material Color Painter}
'{ControlTipText: Apply industry standard colors to selection}
'
'------Buttons------------------------------
' %UI Label lbl_steel Steel Grades:
' %UI Button btn_mild Mild Steel (<210)
' %UI Button btn_hss HSS (210-340)
' %UI Button btn_ahss AHSS (340-590)
' %UI Button btn_uhss UHSS (590-980)
' %UI Button btn_Gpa Giga (980-1200)
' %UI Button btn_HF HotForm (>1200)
' %UI Label lbl_other Other Materials:
' %UI Button btn_Alu Aluminum
' %UI Button btn_Fas Fasteners
' %UI Button btn_cancel Close
'------------------------------------------


Private oprt As Object
Private Const mdlName As String = "MDL_MaterialColors"

Sub ShowMaterialPainter()
    Set oprt = Nothing
'    Dim mapmdl: Set mapmdl = KCL.setBTNmdl(mdlName)
'    Dim mapFunc As Object: Set mapFunc = KCL.InitDic
'        mapFunc.Add "btn_mild", "btn_mild_click"  '这里可以改为mapfunc(BtnName)=BtnName & "_click"
'        mapFunc.Add "btnWrite", "rvme"
'     Set g_Frm = Nothing:  Set g_Frm = KCL.newFrm(mdlName)
'        g_Frm.ShowToolbar mdlName, mapmdl, mapFunc
Set g_Frm = Nothing
    Set g_Frm = KCL.newFrm(mdlName, True):
    g_Frm.Show vbModeless, Nothing, Nothing

Set oprt = KCL.get_inwork_part
 If oprt Is Nothing Then Exit Sub
 
color_mildsteel = Array(169, 169, 169)
        color_HSS = Array(34, 139, 34)
        color_AHSS = Array(255, 215, 0)
    color_UHSS = Array(255, 140, 0)
    color_Gpa = Array(220, 20, 60)
    color_HF = Array(148, 0, 211)
    color_Alu = Array(0, 191, 255)
    color_Fas = Array(139, 69, 19)


    Select Case LCase(g_Frm.BtnClicked)
        Case "btn_mild"
            ApplyColor color_mildsteel
        Case "btn_hss"
            ApplyColor color_HSS
        Case "btn_ahss"
            ApplyColor color_AHSS
        Case "btn_uhss"
            ApplyColor color_UHSS
        Case "btn_gpa"
            ApplyColor color_Gpa
        Case "btn_HF"
            ApplyColor color_HF
        Case "btn_alu"
            ApplyColor color_Alu
        Case "btn_fastener"
            ApplyColor color_Fas
        Case "btn_cancel"
            ' Do nothing, just close
    End Select
 
End Sub
Sub btn_mild_click()
    color_Alu = Array(0, 191, 255)
ApplyColor color_Alu

End Sub

Private Sub ApplyColor(ary As Variant)
    On Error Resume Next
        Dim oSel As Selection
        Set oSel = CATIA.ActiveDocument.Selection
        If oSel.count = 0 Then
            MsgBox "Please select a body or part first.", vbExclamation
            Exit Sub
        End If
        
        R = ary(0): G = ary(1): B = ary(2)
        
        ' Set Real Color (R, G, B, Inheritance=1)
        oSel.VisProperties.SetRealColor R, G, B, 1
        CATIA.StatusBar = "Applied color: "
    On Error GoTo 0
End Sub


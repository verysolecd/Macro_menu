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
' %UI Button btn_giga Giga (980-1200)
' %UI Button btn_hot HotForm (>1200)
' %UI Label lbl_other Other Materials:
' %UI Button btn_alu Aluminum
' %UI Button btn_fastener Fasteners
' %UI Button btn_cancel Close
'------------------------------------------

Option Explicit
Private Const mdlname As String = "MDL_MaterialColors"

Sub ShowMaterialPainter()
    Dim mapmdl: Set mapmdl = KCL.setBTNmdl(mdlname)
    Dim mapFunc As Object: Set mapFunc = KCL.InitDic
        mapFunc.Add "btn_mild", "btn_mild_click"  '这里可以改为mapfunc(BtnName)=BtnName & "_click"
        mapFunc.Add "btnWrite", "rvme"
    
     Set g_Frm = Nothing:  Set g_Frm = KCL.newFrm(mdlname)
        g_Frm.ShowToolbar mdlname, mapmdl, mapFunc
    
'    Select Case ii
'        Case "btn_mild"
'            ApplyColor 169, 169, 169, "Mild Steel"
'        Case "btn_hss"
'            ApplyColor 34, 139, 34, "HSS"
'        Case "btn_ahss"
'            ApplyColor 255, 215, 0, "AHSS"
'        Case "btn_uhss"
'            ApplyColor 255, 140, 0, "UHSS"
'        Case "btn_giga"
'            ApplyColor 220, 20, 60, "GigaPascal"
'        Case "btn_hot"
'            ApplyColor 148, 0, 211, "Hot Formed"
'        Case "btn_alu"
'            ApplyColor 0, 191, 255, "Aluminum"
'        Case "btn_fastener"
'            ApplyColor 139, 69, 19, "Fastener"
'        Case "btn_cancel"
'            ' Do nothing, just close
'    End Select
 
End Sub





Private Sub ApplyColor(R As Long, G As Long, B As Long, Desc As String)
    On Error Resume Next
        Dim oSel As Selection
        Set oSel = CATIA.ActiveDocument.Selection
        If oSel.count = 0 Then
            MsgBox "Please select a body or part first.", vbExclamation
            Exit Sub
        End If
        ' Set Real Color (R, G, B, Inheritance=1)
        oSel.VisProperties.SetRealColor R, G, B, 1
        CATIA.StatusBar = "Applied color: " & Desc
    On Error GoTo 0
End Sub


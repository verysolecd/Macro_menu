Attribute VB_Name = "MDL_MaterialColors"

'{GP:4}
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

Option Explicit
Private Const mdlName As String = "MDL_MaterialColors"

' Main Entry Point
Sub MaterialPainter()
    ' 1. Map all buttons to the SAME handler
    Dim mapFunc As Object: Set mapFunc = KCL.InitDic
    Set mapFunc = setMasterFunc(mdlName)
    
    ' 2. Map Module (Self)
    Dim mapMdl As Object: Set mapMdl = KCL.InitDic
    Dim k
    For Each k In mapFunc.keys
        mapMdl.Add k, mdlName
    Next

    ' 3. Initialize Form with PassButtonName ENABLED
    Set g_Frm = Nothing
    Set g_Frm = KCL.newFrm(mdlName)
    g_Frm.PassButtonName = True ' <--- The Magic Switch
    
    ' 4. Show Toolbar (Modeless)
    g_Frm.ShowToolbar mdlName, mapMdl, mapFunc
End Sub
Function setMasterFunc(ByVal modName As String)
    Set setMasterFunc = Nothing
    Dim ctrllst:    Set ctrllst = KCL.ParseUIConfig(KCL.getbf1stproc(modName))
    Dim map: Set map = KCL.InitDic
    Dim ctrl
    For Each ctrl In ctrllst    '映射BTN名字和对应函数
        Select Case ctrl("Type")
            Case "Forms.CommandButton.1"
                map(ctrl("Name")) = "Action_ClickHandler"
        End Select
    Next
   Set setMasterFunc = map
End Function
' Centralized Handler - Receives Button Name!
Sub Action_ClickHandler(ByVal btnName As String)
    If btnName = "btn_cancel" Then
        Unload g_Frm
        Exit Sub
    End If

'材料 / 强度等级 强度区间（屈服强度）    行业通用推荐配色（RGB） 沃尔沃图中实际配色（RGB）   典型应用部位
'普通低碳钢（Mild steel）    ≤210MPa    浅蓝色（173, 216, 230） 浅灰色（211, 211, 211） 车身蒙皮、后围板等非承力件
'高强度钢（High strength steel） 210-590MPa  深蓝色（0, 0, 205） 浅蓝色（173, 216, 230） 地板横梁、门槛梁外板
'先进高强钢（Very high strength steel）  590-980MPa  黄色（255, 255, 0） 黄色（255, 255, 0） 底板加强件、前后纵梁
'超高强钢（Extra high strength steel）   980-1500MPa 橙色（255, 165, 0） 橙色（255, 165, 0） 车门防撞梁、座椅横梁
'热成型钢（Ultra high strength steel）   1500-2000MPa    红色（255, 0, 0）   红色（255, 0, 0）   A 柱、B 柱、C 柱、门槛梁内板
'铝合金（Aluminium） 150-350MPa（屈服强度）  银色（192, 192, 192）   绿色（0, 255, 0）   前防撞梁、引擎盖骨架、悬挂部件


    ' Define Colors (Preserved from original)
    Dim color_mildsteel, color_HSS, color_AHSS, color_UHSS, color_Gpa, color_HF, color_Alu, color_Fas
    color_mildsteel = Array(169, 169, 169)
    color_HSS = Array(34, 139, 34)
    color_AHSS = Array(255, 215, 0)
    color_UHSS = Array(255, 140, 0)
    color_Gpa = Array(220, 20, 60)
    color_HF = Array(148, 0, 211)
  
    color_Fas = Array(139, 69, 19)
    color_Alu = Array(210, 210, 210)
      color_Alu = Array(160, 160, 160)
    
    Dim mColor As Variant
    Select Case btnName
        Case "btn_mild": mColor = color_mildsteel
        Case "btn_hss": mColor = color_HSS
        Case "btn_ahss": mColor = color_AHSS
        Case "btn_uhss": mColor = color_UHSS
        Case "btn_Gpa": mColor = color_Gpa
        Case "btn_HF": mColor = color_HF
        Case "btn_Alu": mColor = color_Alu
        Case "btn_Fas": mColor = color_Fas
        Case Else: Exit Sub
    End Select
    
    ApplyColor mColor
End Sub

Private Sub ApplyColor(ary As Variant)
    On Error Resume Next
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    
    If oSel.count = 0 Then
        MsgBox "Please select a body or part first.", vbExclamation
        Exit Sub
    End If
    
    Dim r, g, b
    r = ary(0): g = ary(1): b = ary(2)
    
    ' Set Real Color (R, G, B, Inheritance=1)
    oSel.VisProperties.SetRealColor r, g, b, 1
    CATIA.StatusBar = "Applied color RGB(" & r & "," & g & "," & b & ")"
    On Error GoTo 0
End Sub


Sub getcolor()

    Dim r, g, b
 r = CLng(0)
 g = CLng(0)
 b = CLng(0)
 Dim ss
 Set ss = CATIA.ActiveDocument.Selection.VisProperties
 ss.GetRealColor r, g, b
 Dim ary
 ary = Array(r, g, b)
 
 Debug.Print r & "," & g & "," & b
 




End Sub


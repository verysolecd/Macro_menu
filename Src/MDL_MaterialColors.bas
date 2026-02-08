Attribute VB_Name = "MDL_MaterialColors"
'{GP:4}
'{EP:MaterialPainter}
'{Caption:实体上色}
'{ControlTipText: Apply industry standard colors to selection}
'
'------Buttons------------------------------
' %UI Button btn_mild 软钢(<210)    #ADD8E6
' %UI Button btn_hss 高强钢(210-340)  #00BFFF
' %UI Button btn_ahss 先进高强(340-590)  #FFFF00
' %UI Button btn_uhss 超高强(590-980) #FFA500
' %UI Button btn_Gpa Gpa钢 (980-1200) #ff0033
' %UI Button btn_HF 热成型 (>1200) #B22222
' %UI Label bl_steel ----------
' %UI Button btn_Alu1 铝合金(<180)  #90EE90
' %UI Button btn_Alu2 铝合金(180~240)  #8FBC8F
' %UI Button btn_Alu3 铝合金(>240) #228B22
' %UI Button btn_Fas 紧固件      #A52A2A
' %UI Button btn_glue 胶水 #C8A2C8

'≤210MPa       浅蓝色    MS=Array(173,216,230)  #ADD8E6
'210-340MPa    深天蓝     HSS=Array(0,191,255)      #00BFFF
'340-590MPa    黄色      AHSS=Array(255,255,0)    #FFFF00
'590-980MPa   橙色      UHSS=Array(255,165,0)    #FFA500

'980-1200MPa  橙红色   Gpa=Array(255,0,51)    #ff0033
'1200-1600    深粉色      HF=Array(255,20,147)      #FF1493
'＜280MPa      浅绿色    Alu=Array(144,238,144) #90EE90
'180~240      深海洋绿    Alu2=Array(34,139,34)   #8FBC8F
'≥280MPa       深绿色    Alu2=Array(34,139,34)  #228B22

' 紧固件       棕色      Fas=Array(165, 42, 42)     #A52A2A
'Glue          淡紫色    Glue=Arrary(200,160,200)  #C8A2C8

'------------------------------------------
Option Explicit
Private mprt
Private mHSF
Private Const mdlname As String = "MDL_MaterialColors"
' Main Entry Point
Sub MaterialPainter()
  Set mprt = KCL.get_inwork_part
  If mprt Is Nothing Then
        Dim doc
        For Each doc In CATIA.Documents
            If TypeName(doc) = "PartDocument" Then
                Set mprt = doc.part
                Exit For
            End If
        Next
    End If
    If mprt Is Nothing Then Exit Sub
    Set mHSF = mprt.HybridShapeFactory
    Dim mapFunc: Set mapFunc = setMasterFunc(mdlname)
    Dim mapMdl: Set mapMdl = KCL.setBTNmdl(mdlname)
    ' 3. Initialize Form with PassButtonName ENABLED
    Set g_frm = Nothing
    Set g_frm = KCL.newFrm(mdlname)
    g_frm.PassButtonName = True ' <--- The Magic Switch
    ' 4. Show Toolbar (Modeless)
    g_frm.ShowToolbar mdlname, mapMdl, mapFunc
End Sub
Sub Action_ClickHandler(ByVal btnName As String)
    If btnName = "btn_cancel" Then
        Unload g_frm
        Exit Sub
    End If
    Dim map: Set map = btn2case(mdlname)
    Dim mColor As Variant
    If map(btnName) <> "" Then mColor = KCL.ParseBDcolor(map(btnName))
    If IsArray(mColor) Then ApplyColor mColor
End Sub
Private Sub ApplyColor(ary As Variant)
    Dim oSel
    Set oSel = CATIA.ActiveDocument.Selection
    Dim R, G, B, i
    R = ary(0): G = ary(1): B = ary(2)
    If oSel.count = 0 Then
        Set oSel = KCL.Selectmulti("请选择BODY")
    End If
  Dim lst: Set lst = KCL.Initlst
  Dim itm, itp
   For i = 1 To oSel.count
         Set itm = oSel.item(i).Value
         Set itp = Nothing
         Set itp = KCL.GetParent_Of_T(itm, "Body")
         If Not itp Is Nothing Then
            lst.Add itp
         Else
            On Error Resume Next
                Dim itype:  itype = mHSF.GetGeometricalFeatureType(itm)
                Error.Clear
            On Error GoTo 0
         End If
        If itype = 7 Then lst.Add itm
    Next i
oSel.Clear
Set itm = Nothing
For Each itm In lst
    oSel.Add itm
Next
    oSel.VisProperties.SetRealColor R, G, B, 0 '(R, G, B, Inheritance=1)
    oSel.Clear
    On Error GoTo 0
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
Sub getcolor()
Dim R, G, B
 R = CLng(0)
 G = CLng(0)
 B = CLng(0)
 Dim ss
 Set ss = CATIA.ActiveDocument.Selection.VisProperties
 ss.GetRealColor R, G, B
 Dim ary
 ary = Array(R, G, B)
 Debug.Print R & "," & G & "," & B
 End Sub
Function btn2case(ByVal modName As String)
    Set btn2case = Nothing
    Dim ctrllst:    Set ctrllst = KCL.ParseUIConfig(KCL.getbf1stproc(modName))
    Dim map: Set map = KCL.InitDic
    Dim ctrl
    For Each ctrl In ctrllst
        Select Case ctrl("Type")
            Case "Forms.CommandButton.1"
                map(ctrl("Name")) = ctrl("Color")
        End Select
    Next
   Set btn2case = map
End Function

Attribute VB_Name = "MDL_MaterialColors"
'{GP:4}
'{EP:MaterialPainter}
'{Caption:实体上色}
'{ControlTipText: 上色Toolbar}
'------控件清单--------------------------------------------------
'控件格式为 %UI <Type> <Name> <Caption/Text><Color:HEX16>
'------Buttons------------------------------
' %UI Button btn_weld 焊缝颜色 #FFFF00
' %UI Button btn_thread 螺纹孔分颜色
' %UI Label  lb_steel ----------
' %UI Button btn_mild 软钢(<210)    #ADD8E6
' %UI Button btn_hss 高强钢(210-340)  #00BFFF
' %UI Button btn_ahss 先进高强(340-590)  #FFFF80
' %UI Button btn_uhss 超高强(590-980) #FFA500
' %UI Button btn_Gpa Gpa钢 (980-1200) #ff0033
' %UI Button btn_HF 热成型 (>1200) #B22222
' %UI Label  lb_steel ----------
' %UI Button btn_Alu1 铝合金(<180)  #90EE90
' %UI Button btn_Alu2 铝合金(180~240)  #8FBC8F
' %UI Button btn_Alu3 铝合金(>240) #228B22
' %UI Button btn_Fas 紧固件      #A52A2A
' %UI Button btn_glue 胶水 #FF00FF
' %UI Label bl_steel ----------
' %UI Button btn_blue 蓝色 #0000FF


'颜色定义
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




'  catia的basic颜色对应表

'序号    颜色名称 (中文 + 英文)  HEX 色值    RGB 数值
'基础标准纯色
'1   黑色 Black  #000000 0,0,0
'2   白色 White  #FFFFFF 255,255,255
'3   红色 Red    #FF0000 255,0,0
'4   黄色 Yellow #FFFF00 255,255,0
'5   绿色 Green  #00FF00 0,255,0
'6   蓝色 Blue   #0000FF 0,0,255
'7   青色 Cyan   #00FFFF 0,255,255
'8   品红色 Magenta  #FF00FF 255,0,255
'9   橙色 Orange #FF8000 255,128,0
'10  深绿色 Dark Green   #008000 0,128,0
'11  深蓝色 Dark Blue    #0000A0 0,0,160
'12  紫色 Purple #800080 128,0,128
'13  紫罗兰色 Violet #8000FF 128,0,255
'14  酒红色 Burgundy #800000 128,0,0
'浅调常用色
'15  淡黄色 Pale Yellow  #FFFF80 255,255,128
'16  淡绿色 Pale Green   #80FF80 128,255,128
'17  淡蓝色 Pale Blue    #8080FF 128,128,255
'18  浅紫色 Light Purple #FF0080 255,0,128
'19  春绿色 Spring Green #00FF80 0,255,128
'20  黄绿色 Chartreuse   #80FF00 128,255,0
'21  青绿色 Blue Green   #00FF40 0,255,64
'22  碧绿色 Aquamarine   #008040 0,128,64
'23  水鸭青 Teal #008080 0,128,128
'24  宝蓝色 Royal Blue   #0080C0 0,128,192
'25  道奇蓝 Dodger Blue  #0080FF 0,128,255
'26  深岩蓝 Dark Slate Blue  #8080C0 128,128,192
'27  深岩灰 Dark Slate Grey  #80FFFF 128,255,255
'28  亮粉色 Hot Pink #FF80C0 255,128,192
'29  兰花紫 Orchid   #FF80FF 255,128,255
'30  粉橙色 Pink Orange  #FF8040 255,128,64
'31  鲜肉色 Salmon   #FF8080 255,128,128
'32  梅子色 Plum #800040 128,0,64
'33  栗色 Maroon #804040 128,64,64
'34  蓝灰色 Blue Grey    #333366 51,51,102
'35  28% 灰色 Grey-28%   #C1C4C0 193,196,192
'小众特殊杂色
'36  沙黄色 Sandy Yellow #FFBE47 255,190,71
'37  金黄色 Golden Yellow    #FABE47 250,190,71
'38  橙肉色 Orange Salmon    #F2A257 242,162,87
'39  粉肉色 Pink Salmon  #EA8466 234,132,102
'40  浅淡紫色 Light Lavender #C4B3D1 196,179,209
'41  深淡紫色 Dark Lavender  #9993BF 153,147,191
'42  岩蓝色 Slate Blue   #83AAD6 131,170,214
'43  天蓝色 Sky Blue #81C0E8 129,192,232
'44  海绿色 Sea Green    #94C9BF 148,201,191
'45  浅海绿 Light Sea Green  #AED19B 174,209,155
'46  浅卡其绿 Light Khaki Green  #BFC990 191,205,144
'47  深灰绿 Dark Grey Green  #7EA297 126,162,151
'48  沙棕色 Sandy Brown  #D3B27D 211,178,125















'------------------------------------------
Option Explicit
Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private a_Prt     ' 当前任意一个零件对象
Private m_sel         As Selection      ' 选择集对象
Private mHSF
Private Const TYPE_SWEEP As Long = 7
Private Type Threadspec
    MinDia As Double                             ' 最小直径 (包含)
    MaxDia As Double                             ' 最大直径 (包含)
    R As Integer                                 ' 红色通道 (0-255)
    G As Integer                                 ' 绿色通道 (0-255)
    B As Integer                                 ' 蓝色通道 (0-255)
End Type
Private oSpec() As Threadspec
Private Colormap


Private mWD As Cls_DynaWD
Private funcMap
Private Const mdlname As String = "MDL_MaterialColors"
' Main Entry Point
Sub MaterialPainter()
    If Not CanExecute("Productdocument,PartDocument") Then Exit Sub
    On Error Resume Next
        Set m_Doc = CATIA.ActiveDocument
        Dim adoc
        For Each adoc In CATIA.Documents
            If TypeName(adoc) = "PartDocument" Then: Set a_Prt = adoc.part: Exit For
        Next
        Err.Clear
    On Error GoTo 0
    If IsNothing(a_Prt) Then: MsgBox "No part found": Exit Sub
    Set mWD = New Cls_DynaWD
    initMaps mdlname  '初始化按钮和颜色map
    mWD.PassBtnName = True ' <--- The Magic Switch
    mWD.ShowToolbar mdlname, , funcMap   ' 4. Show Toolbar (Modeless) — modMap 自动构建, 仅传自定义 macMap
End Sub
Sub Clickhandler(ByVal btnName As String)
    If btnName = "btn_cancel" Then: Set mWD = Nothing: Exit Sub
    '先检查传入按钮的名字具备对应颜色
    Dim mcolor
    If funcMap(btnName) <> "" Then mcolor = KCL.ParseBDcolor(Colormap(btnName))
    Select Case btnName
        Case "btn_weld"
            SetWeldYellow
        Case "btn_tsshread"
            SetThreadColor
        Case Else
              If IsArray(mcolor) Then ApplyColor2Body mcolor
    End Select
End Sub
Private Sub ApplyColor2Body(ary As Variant)
    Set mHSF = a_Prt.HybridShapeFactory
    Dim osel
    Set osel = CATIA.ActiveDocument.Selection
    Dim R, G, B, i
    R = ary(0): G = ary(1): B = ary(2)
    If osel.count = 0 Then
        Set osel = KCL.Selectmulti("请选择BODY")
    End If
  Dim lst: Set lst = KCL.Initlst
  Dim itm, itp
   For i = 1 To osel.count
         Set itm = osel.item(i).Value
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
osel.Clear
Set itm = Nothing
For Each itm In lst
    osel.Add itm
Next
    osel.VisProperties.SetRealColor R, G, B, 0 '(R, G, B, Inheritance=1)
    osel.Clear
    On Error GoTo 0
End Sub

Sub getcolor()
    Dim R, G, B
        R = CLng(0)
        G = CLng(0)
        B = CLng(0)
 Dim ss: Set ss = CATIA.ActiveDocument.Selection.VisProperties
    ss.GetRealColor R, G, B
 Dim ary: ary = Array(R, G, B)
 Debug.Print "RGB颜色" & R & "," & G & "," & B
End Sub
Private Sub initMaps(ByVal modName As String)
    Set funcMap = KCL.InitDic
    Set Colormap = KCL.InitDic
    Dim ctrllst: Set ctrllst = KCL.ParseUIConfig(KCL.getbf1stproc(modName))
    Dim ctrl
    For Each ctrl In ctrllst
        If ctrl("Type") = "Forms.CommandButton.1" Then
            Dim n As String: n = ctrl("Name")
            funcMap(n) = "Clickhandler"
            Colormap(n) = ctrl("Color")
        End If
    Next
End Sub
Sub SetWeldYellow()
   If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
   If m_prt Is Nothing Then Exit Sub
    Dim c, Color, i
    Dim HSF:  Set HSF = m_prt.HybridShapeFactory
    Dim sweeps: Set sweeps = KCL.Initlst
        m_sel.Clear
        CATIA.HSOSynchronized = False
            Set m_sel = KCL.SelectQuery("(.'Volume geometry'+.Surface& Type!=Plane)& Color!=Yellow")
            Color = Array(255, 255, 0)
            m_sel.VisProperties.SetRealColor Color(0), Color(1), Color(2), 0 '(R, G, B, Inheritance=1)
            m_sel.Clear
        CATIA.HSOSynchronized = True
End Sub

Sub SetThreadColor()
    On Error GoTo ErrorHandler
    Dim oCatia As Object
    Set oCatia = CATIA
    If oCatia.Documents.count = 0 Then
        MsgBox "请先打开一个 Part 或 Product 文档。", vbExclamation
        Exit Sub
    End If
    Dim oDoc As Document
    Set oDoc = oCatia.ActiveDocument
    Set m_sel = oDoc.Selection
    oCatia.DisplayFileAlerts = False
    oCatia.RefreshDisplay = False
    
    Call initThreadSpec
    Select Case TypeName(oDoc)
    Case "PartDocument"
        Call ProcessPart(oDoc.part)
    Case "ProductDocument"
        Call ProcessProduct(oDoc.Product)
    Case Else
        MsgBox "此宏仅能在 Part 或 Product 环境下运行。", vbExclamation
    End Select
Cleanup:
    oCatia.RefreshDisplay = True
    oCatia.DisplayFileAlerts = True
    MsgBox "螺纹染色处理完成！", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "运行期间发生意外错误: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub initThreadSpec()
    ReDim oSpec(3)
    With oSpec(0) '  M4 (D=4.0, 范围 3.6 ~ 4.4) -> 黄色 (Yellow)
        .MinDia = 3.6: .MaxDia = 4.4
        .R = 255: .G = 255: .B = 0
    End With
    With oSpec(1) 'M5 (D=5.0, 范围 4.6 ~ 5.4) -> 紫色 (Purple)
        .MinDia = 4.6: .MaxDia = 5.4
        .R = 255: .G = 0: .B = 255
    End With
    With oSpec(2) 'M6 (D=6.0, 范围 5.6 ~ 6.4) -> 绿色 (Green)
        .MinDia = 5.6: .MaxDia = 6.4
        .R = 0: .G = 255: .B = 0
    End With
    With oSpec(3) 'M8及以上 (D=8.0+, 范围 7.6 ~ 100) -> 蓝色 (Blue)
        .MinDia = 7.6: .MaxDia = 100
        .R = 0: .G = 0: .B = 255
    End With
End Sub
Private Sub ProcessProduct(ByVal oProd As Product)
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim childCount As Integer
    childCount = oProd.Products.count
    If childCount = 0 Then
        Dim oPart As part
        Set oPart = TryGetPartFromProduct(oProd)
        If Not oPart Is Nothing Then
            On Error GoTo ErrorHandler
            oProd.ApplyWorkMode 2  ' 2 = DESIGN_MODE (应用设计模式以加载零件完整特征)
            Call ProcessPart(oPart)
        End If
    Else
        For i = 1 To childCount
            Call ProcessProduct(oProd.Products.item(i))
        Next i
    End If
    Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub ProcessPart(ByVal oPart As part)
    Dim oBody As body
    Dim oShape As Object
    Dim i As Integer, j As Integer
    Dim dThreadDia As Double
    Dim bIsThread As Boolean
    Dim R As Integer, G As Integer, B As Integer
    ' ── 改动：按颜色分组存储特征，key = "R,G,B"
    Dim colorGroups As Object
    Set colorGroups = CreateObject("Scripting.Dictionary")
    For i = 1 To oPart.bodies.count
        Set oBody = oPart.bodies.item(i)
        For j = 1 To oBody.Shapes.count
            Set oShape = oBody.Shapes.item(j)
            bIsThread = False
            dThreadDia = 0#
            Select Case TypeName(oShape)
            Case "Hole"
                If oShape.ThreadingMode = 0 Then
                    If TryGetHoleThreadDiameter(oShape, dThreadDia) Then
                        bIsThread = True
                    End If
                End If
            Case "Thread"
                If TryGetThreadDiameter(oShape, dThreadDia) Then
                    bIsThread = True
                End If
            End Select
            ' ── 改动：不立刻上色，按颜色 key 归组
            If bIsThread Then
                If GetColorByDia(dThreadDia, R, G, B) Then
                    Dim colorKey As String
                    colorKey = R & "," & G & "," & B
                    If Not colorGroups.Exists(colorKey) Then
                        Dim lst: Set lst = CreateObject("System.Collections.ArrayList")
                      colorGroups.Add colorKey, lst
                    End If
                    colorGroups(colorKey).Add oShape
                End If
            End If
        Next j
    Next i
    ' ── 改动：每种颜色只调用一次 SetRealColor（批量）
    Dim key As Variant
    Dim rgb() As String
    Dim shapeList As Object
    Dim shp As Object
    For Each key In colorGroups.keys
        rgb = Split(key, ",")
        Set shapeList = colorGroups(key)
        m_sel.Clear
        For Each shp In shapeList
            On Error Resume Next
            m_sel.Add shp
            On Error GoTo 0
        Next
        If m_sel.count > 0 Then
            m_sel.VisProperties.SetRealColor CLng(rgb(0)), CLng(rgb(1)), CLng(rgb(2)), 1
        End If
    Next key
    m_sel.Clear
End Sub

Private Function TryGetPartFromProduct(ByVal oProd As Product) As part
    On Error GoTo Fail
    Set TryGetPartFromProduct = oProd.ReferenceProduct.Parent.part
    Exit Function
Fail:
    Set TryGetPartFromProduct = Nothing
    Err.Clear
End Function
Private Function TryGetHoleThreadDiameter(ByVal oHole As Object, ByRef outDiameter As Double) As Boolean
    On Error GoTo Fail
        outDiameter = oHole.ThreadDiameter.Value
        TryGetHoleThreadDiameter = True
    Exit Function
Fail:
    TryGetHoleThreadDiameter = False
    Err.Clear
End Function
Private Function TryGetThreadDiameter(ByVal oThread As Object, ByRef outDiameter As Double) As Boolean
    On Error GoTo Fail
    outDiameter = oThread.Diameter
    TryGetThreadDiameter = True
    Exit Function
Fail:
    TryGetThreadDiameter = False
    Err.Clear
End Function
Private Function GetColorByDia(ByVal dDia As Double, ByRef outR As Integer, ByRef outG As Integer, ByRef outB As Integer) As Boolean
    Dim k As Integer
    GetColorByDia = False
    For k = LBound(oSpec) To UBound(oSpec)
        If dDia >= oSpec(k).MinDia And dDia <= oSpec(k).MaxDia Then
            outR = oSpec(k).R
            outG = oSpec(k).G
            outB = oSpec(k).B
            GetColorByDia = True
            Exit Function
        End If
    Next k
End Function


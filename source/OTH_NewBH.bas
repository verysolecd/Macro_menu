Attribute VB_Name = "OTH_NewBH"
'Attribute VB_Name = "m30_BH"
'{GP:6}
'{Ep:CATMain}
'{Caption:新电池箱体}
'{ControlTipText:新建一个电池箱体的结构树}
'{BackColor:}
Private Tree As Variant
Private prj
Sub CATMain()
        Dim oprd, Tree, odoc, rootprd, cover, house, ref
          Dim imsg
          imsg = "请输入你的项目名称"
        prj = GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
        Set Tree = KCL.InitDic(vbTextCompare)
        Call iniTree(Tree)
        '====创建参考和包络===
    For i = 0 To 18
        Select Case i
            Case 0      '创建根产品
                Set odoc = CATIA.Documents.Add("Product")
                Set rootprd = odoc.Product
                Set oprd = rootprd
            Case 1
                Set oprd = rootprd.Products.AddNewComponent("Product", "")
            Case 2
                 Set oprd = rootprd.Products.AddNewComponent("Product", "")
            Case 3      '创建上箱体组件
                Set cover = rootprd.Products.AddNewComponent("Product", "")
                Set oprd = cover
           Case 4     '创建上箱体
                Set oprd = cover.Products.AddNewComponent("Part", "")
           Case 5     '创建下箱体
                Set house = rootprd.Products.AddNewComponent("Product", "")
                Set oprd = house
            Case 6 '创建参考件
                Set ref = house.Products.AddNewComponent("Part", "")
                Set oprd = ref
            Case 12, 13, 14 '创建零件
                Set oprd = house.Products.AddNewComponent("Part", "")
                oprd.Name = Tree(i)(4)
            Case 15, 16:
               Set oprd = house.Products.AddNewProduct("")
            Case 17
               Set oprd = rootprd.Products.AddNewComponent("Product", "")
            Case 18
                Set fast = rootprd.Products.AddNewComponent("Product", "")
                Set oprd = fast
        Case Else
                Set oprd = house.Products.AddNewComponent("Product", "")
        End Select
          Call newPn(oprd, Tree(i))
          Set oprd = Nothing
    Next
        '===新增component
     ' Set product4 = products1.AddNewProduct("")
    ' 新增产品= oprd.products.AddNewComponent("Part", "")
    Dim osel
    Set osel = CATIA.ActiveDocument.Selection
    osel.Clear
    osel.Add ref
    osel.Copy
    osel.Clear
    Dim otp
    Set otp = CATIA.ActiveDocument.Selection
    otp.Clear
    otp.Add fast
    otp.Paste
    
    Call initme
    
    
End Sub

Sub iniTree(Tree)
    Tree(0) = Array(0, "_Prj_Housing_Asm", "Project Housing Asm", "箱体组件", "Housing Asm")
    Tree(1) = Array(0, "_Pack", "Pack system", "整包方案", "Pack system")
    Tree(2) = Array(0, "_Packaging", "packaging", "包络定义", "packaging")
    Tree(3) = Array(0, "_000", "Upper Housing Asm", "上箱体总成", "Upper Housing Asm")
    Tree(4) = Array(0, "_001", "Upper Housing", "上箱体", "Upper Housing")
    Tree(5) = Array(0, "_1000", "Lower Housing Asm", "下箱体总成", "Lower Housing Asm")
    Tree(6) = Array(0, "_ref", "Ref", "参考", "Ref")
    Tree(7) = Array(0, "_1100", "Frames", "框架组件", "Frames")
    Tree(8) = Array(0, "_1200", "Members", "梁组件", "Members")
    Tree(9) = Array(0, "_1300", "Brkts", "支架组件", "Brkts")
    Tree(10) = Array(0, "_1400", "Bottom components", "底部组件", "Bottom components")
    Tree(11) = Array(0, "_1500", "Cooling system", "液冷组件", "Cooling system")
    Tree(12) = Array(0, "_2001", "Weldings Seams", "焊缝", "Weldings Seams")
    Tree(13) = Array(0, "_2002", "SPot Welding", "点焊", "Spot Welding")
    Tree(14) = Array(0, "_2003", "Adhesive", "胶水", "adhesive")
    
    Tree(15) = Array(0, "_4000", "Grou_fasteners", "紧固件组合", "Group_Fastener.1")
    Tree(16) = Array(0, "_5000", "others", "其他组件", "others")
    
    Tree(17) = Array(0, "_Abandon", "Abandoned", "废案", "Abandoned")
    Tree(18) = Array(0, "_Patterns", "Fasteners", "紧固件阵列", "Fasteners Pattern")
    
End Sub

Sub newPn(oprd, arr)
    oprd.Name = arr(4)
    oprd.PartNumber = prj & arr(1)
    oprd.nomenclature = arr(2)
    oprd.definition = arr(3)
    On Error Resume Next
    oprd.Name = arr(4)
    On Error GoTo 0
    oprd.Name = arr(4)
    oprd.Update
End Sub

Public Function GetInput(msg) As String
    Dim UserInput As String
    UserInput = InputBox(msg, "输入提示")
    
    ' 如果用户没有输入或点击取消，则返回默认值"XX"
    If UserInput = "" Or UserInput = "0" Then
        GetInput = ""
    Else
        GetInput = UserInput
    End If
End Function

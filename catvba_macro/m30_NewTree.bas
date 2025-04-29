Attribute VB_Name = "m30_NewTree"
'Attribute VB_Name = "m30_NewTree"
'{GP:3}
'{Ep:CATMain}
'{Caption:新电池箱体}
'{ControlTipText:新建一个电池箱体的结构树}
'{BackColor:}
Private Tree As Variant
Private prj
Sub CATMain()
        Dim oprd, Tree, oDoc, rootPrd, cover, house, ref
          Dim imsg
          imsg = "请输入你的项目名称"
        prj = KCL.GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
        Set Tree = KCL.InitDic(vbTextCompare)
        Call iniTree(Tree)
        '====创建参考和包络===
    For i = 0 To 17
        Select Case i
            Case 0      '创建根产品
                Set oDoc = CATIA.Documents.Add("Product")
                Set rootPrd = oDoc.Product
                Set oprd = rootPrd
             Case 1
                Set oprd = rootPrd.Products.AddNewComponent("Product", "")
            Case 2
            Case 3      '创建上箱体组件
                Set cover = rootPrd.Products.AddNewComponent("Product", "")
                Set oprd = cover
           Case 4     '创建上箱体
                Set oprd = cover.Products.AddNewComponent("Part", "")
           Case 5     '创建下箱体
                Set house = rootPrd.Products.AddNewComponent("Product", "")
                Set oprd = house
            Case 6   '创建参考件
                Set ref = house.Products.AddNewComponent("Part", "")
                Set oprd = ref
            Case 14, 15:
               Set oprd = house.Products.AddNewProduct("")
            Case 16
               Set oprd = rootPrd.Products.AddNewComponent("Product", "")
            Case 17
                Set fast = rootPrd.Products.AddNewComponent("Product", "")
                Set oprd = fast
            Case Else
                Set oprd = house.Products.AddNewComponent("Product", "")
        End Select
          Call newPn(oprd, Tree(i))
    Next
        '===新增component
        ' Set product4 = products1.AddNewProduct("")
    ' Set product3 = oprd.products.AddNewComponent("Part", "")
    Dim osel
    Set osel = CATIA.ActiveDocument.Selection
    osel.Clear
    osel.Add ref
    osel.Copy
    Dim otp
    Set otp = CATIA.ActiveDocument.Selection
    otp.Clear
    otp.Add fast
    otp.Paste
End Sub
Sub iniTree(Tree)
    Tree(0) = Array(0, "_Prj_Housing_Asm", "Project Housing Asm", "箱体组件", "Housing Asm")
    Tree(1) = Array(0, "_Pack", "Pack system", "整包方案", "Pack system")
    Tree(2) = Array(0, "_Packaging", "packaging", "包络定义", "packaging")
    
    
    Tree(3) = Array(0, "_0000", "Upper Housing Asm", "上箱体总成", "Upper Housing Asm")
    Tree(4) = Array(0, "_0001", "Upper Housing", "上箱体", "Upper Housing")
    
    
    
    Tree(5) = Array(0, "_1000", "Lower Housing Asm", "下箱体总成", "Lower Housing Asm")
    Tree(6) = Array(0, "_ref", "Ref", "参考", "Ref")
    Tree(7) = Array(0, "_1100", "Sealing components", "密封组件", "Sealing components")
    Tree(8) = Array(0, "_1200", "Frames", "框架组件", "Frames")
    Tree(9) = Array(0, "_1300", "Members", "梁组件", "Members")
    Tree(10) = Array(0, "_1400", "Bottom components", "底部组件", "Bottom components")
    Tree(11) = Array(0, "_1900", "Cooling system", "液冷组件", "Cooling system")
    Tree(12) = Array(0, "_2000", "Weldings", "焊接信息", "weldings")
    Tree(13) = Array(0, "_3000", "Adhesive", "胶水组件", "adhesive")
    Tree(14) = Array(0, "_4000", "Grou_fasteners", "紧固件组合", "Group_Fastener.1")
    Tree(15) = Array(0, "_5000", "others", "其他组件", "others")
    
    
    Tree(16) = Array(0, "_Abandon", "Abandoned", "废案", "Abandoned")
    Tree(17) = Array(0, "_Patterns", "Fasteners", "紧固件阵列", "Fasteners Pattern")
End Sub
Sub newPn(oprd, arr)
    oprd.PartNumber = prj & arr(1)
    oprd.nomenclature = arr(2)
    oprd.definition = arr(3)
    On Error Resume Next
    oprd.Name = arr(4)
    On Error GoTo 0
End Sub

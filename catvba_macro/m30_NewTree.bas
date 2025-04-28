Attribute VB_Name = "m30_NewTree"
'Attribute VB_Name = "m30_NewTree"
'{GP:3}
'{Ep:CATMain}
'{Caption:�µ������}
'{ControlTipText:�½�һ���������Ľṹ��}
'{BackColor:}
Private Tree As Variant
Private prj
Sub CATMain()
        Dim oprd, Tree, oDoc, rootPrd, cover, house, ref
          Dim imsg
          imsg = "�����������Ŀ����"
        prj = KCL.GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
        Set Tree = KCL.InitDic(vbTextCompare)
        Call iniTree(Tree)
        '====�����ο��Ͱ���===
    For i = 0 To 17
        Select Case i
            Case 0      '��������Ʒ
                Set oDoc = CATIA.Documents.Add("Product")
                Set rootPrd = oDoc.Product
                Set oprd = rootPrd
             Case 1
                Set oprd = rootPrd.Products.AddNewComponent("Product", "")
            Case 2
            Case 3      '�������������
                Set cover = rootPrd.Products.AddNewComponent("Product", "")
                Set oprd = cover
           Case 4     '����������
                Set oprd = cover.Products.AddNewComponent("Part", "")
           Case 5     '����������
                Set house = rootPrd.Products.AddNewComponent("Product", "")
                Set oprd = house
            Case 6   '�����ο���
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
        '===����component
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
    Tree(0) = Array(0, "_Prj_Housing_Asm", "Project Housing Asm", "�������", "Housing Asm")
    Tree(1) = Array(0, "_Pack", "Pack system", "��������", "Pack system")
    Tree(2) = Array(0, "_Packaging", "packaging", "���綨��", "packaging")
    
    
    Tree(3) = Array(0, "_0000", "Upper Housing Asm", "�������ܳ�", "Upper Housing Asm")
    Tree(4) = Array(0, "_0001", "Upper Housing", "������", "Upper Housing")
    
    
    
    Tree(5) = Array(0, "_1000", "Lower Housing Asm", "�������ܳ�", "Lower Housing Asm")
    Tree(6) = Array(0, "_ref", "Ref", "�ο�", "Ref")
    Tree(7) = Array(0, "_1100", "Sealing components", "�ܷ����", "Sealing components")
    Tree(8) = Array(0, "_1200", "Frames", "������", "Frames")
    Tree(9) = Array(0, "_1300", "Members", "�����", "Members")
    Tree(10) = Array(0, "_1400", "Bottom components", "�ײ����", "Bottom components")
    Tree(11) = Array(0, "_1900", "Cooling system", "Һ�����", "Cooling system")
    Tree(12) = Array(0, "_2000", "Weldings", "������Ϣ", "weldings")
    Tree(13) = Array(0, "_3000", "Adhesive", "��ˮ���", "adhesive")
    Tree(14) = Array(0, "_4000", "Grou_fasteners", "���̼����", "Group_Fastener.1")
    Tree(15) = Array(0, "_5000", "others", "�������", "others")
    
    
    Tree(16) = Array(0, "_Abandon", "Abandoned", "�ϰ�", "Abandoned")
    Tree(17) = Array(0, "_Patterns", "Fasteners", "���̼�����", "Fasteners Pattern")
End Sub
Sub newPn(oprd, arr)
    oprd.PartNumber = prj & arr(1)
    oprd.nomenclature = arr(2)
    oprd.definition = arr(3)
    On Error Resume Next
    oprd.Name = arr(4)
    On Error GoTo 0
End Sub

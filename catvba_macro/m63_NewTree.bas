Attribute VB_Name = "m63_NewTree"
'Attribute VB_Name = "m30_NewTree"
'{GP:6}
'{Ep:CATMain}
'{Caption:�µ������}
'{ControlTipText:�½�һ���������Ľṹ��}
'{BackColor:}
Private Tree As Variant
Private prj
Sub CATMain()
        Dim oprd, Tree, oDoc, rootprd, cover, house, ref
          Dim imsg
          imsg = "�����������Ŀ����"
        prj = GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
        Set Tree = KCL.InitDic(vbTextCompare)
        Call iniTree(Tree)
        '====�����ο��Ͱ���===
    For i = 0 To 18
        Select Case i
            Case 0      '��������Ʒ
                Set oDoc = CATIA.Documents.Add("Product")
                Set rootprd = oDoc.product
                Set oprd = rootprd
            Case 1
                Set oprd = rootprd.Products.AddNewComponent("Product", "")
            Case 2
                 Set oprd = rootprd.Products.AddNewComponent("Product", "")
            Case 3      '�������������
                Set cover = rootprd.Products.AddNewComponent("Product", "")
                Set oprd = cover
           Case 4     '����������
                Set oprd = cover.Products.AddNewComponent("Part", "")
           Case 5     '����������
                Set house = rootprd.Products.AddNewComponent("Product", "")
                Set oprd = house
            Case 6 '�����ο���
                Set ref = house.Products.AddNewComponent("Part", "")
                Set oprd = ref
            Case 12, 13, 14 '�������
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
        '===����component
     ' Set product4 = products1.AddNewProduct("")
    ' ������Ʒ= oprd.products.AddNewComponent("Part", "")
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
End Sub

Sub iniTree(Tree)
    Tree(0) = Array(0, "_Prj_Housing_Asm", "Project Housing Asm", "�������", "Housing Asm")
    Tree(1) = Array(0, "_Pack", "Pack system", "��������", "Pack system")
    Tree(2) = Array(0, "_Packaging", "packaging", "���綨��", "packaging")
    Tree(3) = Array(0, "_000", "Upper Housing Asm", "�������ܳ�", "Upper Housing Asm")
    Tree(4) = Array(0, "_001", "Upper Housing", "������", "Upper Housing")
    Tree(5) = Array(0, "_1000", "Lower Housing Asm", "�������ܳ�", "Lower Housing Asm")
    Tree(6) = Array(0, "_ref", "Ref", "�ο�", "Ref")
    Tree(7) = Array(0, "_1100", "Frames", "������", "Frames")
    Tree(8) = Array(0, "_1200", "Members", "�����", "Members")
    Tree(9) = Array(0, "_1300", "Brkts", "֧�����", "Brkts")
    Tree(10) = Array(0, "_1400", "Bottom components", "�ײ����", "Bottom components")
    Tree(11) = Array(0, "_1500", "Cooling system", "Һ�����", "Cooling system")
    Tree(12) = Array(0, "_2001", "Weldings Seams", "����", "Weldings Seams")
    Tree(13) = Array(0, "_2002", "SPot Welding", "�㺸", "Spot Welding")
    Tree(14) = Array(0, "_2003", "Adhesive", "��ˮ", "adhesive")
    
    Tree(15) = Array(0, "_4000", "Grou_fasteners", "���̼����", "Group_Fastener.1")
    Tree(16) = Array(0, "_5000", "others", "�������", "others")
    
    Tree(17) = Array(0, "_Abandon", "Abandoned", "�ϰ�", "Abandoned")
    Tree(18) = Array(0, "_Patterns", "Fasteners", "���̼�����", "Fasteners Pattern")
    
End Sub

Sub newPn(oprd, Arr)
    oprd.Name = Arr(4)
    oprd.PartNumber = prj & Arr(1)
    oprd.nomenclature = Arr(2)
    oprd.definition = Arr(3)
    On Error Resume Next
    oprd.Name = Arr(4)
    On Error GoTo 0
    oprd.Name = Arr(4)
    oprd.Update
End Sub

Public Function GetInput(msg) As String
    Dim UserInput As String
    UserInput = InputBox(msg, "������ʾ")
    
    ' ����û�û���������ȡ�����򷵻�Ĭ��ֵ"XX"
    If UserInput = "" Or UserInput = "0" Then
        GetInput = ""
    Else
        GetInput = UserInput
    End If
End Function

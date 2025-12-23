Attribute VB_Name = "OTH_NewBH"
'Attribute VB_Name = "OTH_NewBH"
'{GP:6}
'{Ep:CATMain}
'{Caption:新电池箱体}
'{ControlTipText:新建一个电池箱体的结构树}
'{BackColor:}
'======零件号信息
' %info P,_Prj_Housing_Asm, Project_HousingAsm,箱体组件,HousingAsm
      ' %info P,_Pack,Pack_system,整包方案,Pack_system
      ' %info P,_Packaging, packaging,包络定义,packaging
      ' %info P,_000, Upper_Housing_Asm,上箱体总成,Upper_Housing_Asm
            '%info t,_001, Upper_Housing, 上箱体, Upper_Housing
      ' %info P,_1000,Lower_Housing_Asm,下箱体总成,Lower_Housing_Asm
           ' %info t,_ref, Ref,参考,Ref
           ' %info t,_1100,Frames,框架组件,Frames
           ' %info t,_1200,Brkts,支架类组件,Brkts
           ' %info t,_1300,Cooling_system,液冷组件,Cooling_system
           ' %info t,_1400,Bottom_components,底部组件,Bottom_components
           ' %info t,_2001,Welding_Seams, 焊缝,Welding_Seams
           ' %info t,_2002,SPot_Welding,点焊,Spot_Welding
           ' %info t,_2003,Adhesive,胶水,adhesive
           ' %info c,_4000,Grou_fasteners,紧固件组合,Group_Fastener
           ' %info t,_5000,others,其他组件,others
      ' %info c,_Abandon,Abandoned,废案,Abandoned
      ' %info c,_Patterns,Fasteners,紧固件阵列,Fasteners_Pattern
Private prj
Sub CATMain()
    prj = KCL.GetInput("请输入项目名称"): If prj = "" Then Exit Sub
    Dim Tree As Object: Set Tree = ParsePn(getDecCode())
    Dim PStack As Object: Set PStack = KCL.InitDic
    Dim k, oPrd As Object, ref As Object, fast As Object
    
    For Each k In Tree.keys
        Set oPrd = AddNode(PStack, Tree(k))
        If InStr(1, k, "_ref", 1) > 0 Then Set ref = oPrd
        If InStr(1, k, "_Patterns", 1) > 0 Then Set fast = oPrd
    Next
    
    If Not (ref Is Nothing Or fast Is Nothing) Then
        CATIA.ActiveDocument.Selection.Add ref: CATIA.ActiveDocument.Selection.Copy
        CATIA.ActiveDocument.Selection.Clear: CATIA.ActiveDocument.Selection.Add fast
        CATIA.ActiveDocument.Selection.Paste: CATIA.ActiveDocument.Selection.Clear
    End If
    If PStack.Exists(1) Then Call recurInitPrd(PStack(1))
End Sub

Function AddNode(PStack, D)
    Dim L%: L = IIf(D.Exists("Level"), CInt(D("Level")), 1)
    Dim oPrd, par, TP$: TP = "Product"
    If L < 1 Then L = 1
    If L = 1 Then
        Set oPrd = CATIA.Documents.Add("Product").Product:  Set PStack(1) = oPrd
    Else
        Set par = PStack(IIf(PStack.Exists(L - 1), L - 1, 1))
        If D.Exists("Type") Then
            If UCase(Trim(D("Type"))) = "T" Then TP = "Part"
            If UCase(Trim(D("Type"))) = "C" Then TP = "Component"
        End If
        If TP = "Component" Then Set oPrd = par.Products.AddNewProduct("") Else Set oPrd = par.Products.AddNewComponent(TP, "")
       Set PStack(L) = oPrd
    End If
    
    On Error Resume Next
    oPrd.Name = D("Name")
    With oPrd.ReferenceProduct
        .PartNumber = prj & D("PartNumber"): .Nomenclature = D("Nomenclature"): .Definition = D("Definition")
    End With
    oPrd.Update
    Set AddNode = oPrd
End Function
Function getDecCode()
    On Error Resume Next
    Dim M As Object: Set M = KCL.GetApc().ExecutingProject.VBProject.VBE.Activecodepane.codemodule
    If Not M Is Nothing Then If M.CountOfDeclarationLines > 0 Then getDecCode = M.Lines(1, M.CountOfDeclarationLines)
End Function

Private Function ParsePn(C$) As Object
    Dim RE As Object, M, lst, curL%, H(20) As Integer, curI%
    Set RE = CreateObject("VBScript.RegExp"): Set lst = KCL.InitDic(1)
    RE.Global = True: RE.MultiLine = True: RE.Pattern = "^(\s*)'\s*%info\s+([^,]*),+([^,]*),+([^,]*),+([^,]*),+([^,\r\n]*).*$"
    If RE.TEST(C) Then
        H(0) = -1: H(1) = 0
        For Each M In RE.Execute(C)
            curI = Len(M.SubMatches(0))
            curL = GetLev(curI, curL, H)
            Dim D: Set D = KCL.InitDic(1)
            D.Add "Level", curL: D.Add "Type", Trim(M.SubMatches(1)): D.Add "PartNumber", Trim(M.SubMatches(2))
            D.Add "Nomenclature", Trim(M.SubMatches(3)): D.Add "Definition", Trim(M.SubMatches(4)): D.Add "Name", Trim(M.SubMatches(5))
            lst.Add D("PartNumber"), D
        Next
    End If
    Set ParsePn = lst
End Function

Private Function GetLev(ByVal I As Integer, ByVal L As Integer, ByRef H() As Integer) As Integer
    If L = 0 Or I > H(L) Then
        L = L + 1: If L > UBound(H) Then L = UBound(H)
        H(L) = I
    Else
        While L > 1 And H(L) > I: L = L - 1: Wend
    End If
    GetLev = L
End Function



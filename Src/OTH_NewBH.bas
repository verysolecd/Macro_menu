Attribute VB_Name = "OTH_NewBH"
'Attribute VB_Name = "OTH_NewBH"
'{GP:6}
'{Ep:CATMain}
'{Caption:新电池箱体}
'{ControlTipText:新建一个电池箱体的结构树}
'{BackColor:}
'======零件号信息
' %info P,_0000, _Housing_Asm,箱体组件,Housing_Asm
      ' %info P,_Pack,Pack_system,整包方案,Pack_system
      ' %info P,_Packaging, packaging,包络定义,packaging
      ' %info P,_2000, Upper_Housing_Asm,上箱体总成,Upper_Housing_Asm
            '%info t,_001, Upper_Housing, 上箱体, Upper_Housing
      ' %info P,_1000,Lower_Housing_Asm,下箱体总成,Lower_Housing_Asm
           ' %info t,_ref, Ref,参考,Ref
           ' %info P,_G100,Frames,框架组件,Frames
                ' %info T,_101,FRONT FRAME,前边框,FRONT FRAME
                ' %info T,_102,REAR FRAME,后边框,REAR FRAME
                ' %info T,_103,RH FRAME,右边框,RH FRAME
                ' %info T,_104,LH FRAME,左边框,LH FRAME
                ' %info T,_111,Xmember 1,横梁1,Xmember 1
                ' %info T,_112,Xmember 2,横梁2,Xmember 2
                ' %info T,_113,Xmember 3,横梁3,Xmember 3
           ' %info C,_G200,Brkts,支架类组件,Brkts
           ' %info P,_G300,Cooling_system,液冷组件,Cooling_system
           ' %info P,_G400,Bottom_components,底部组件,Bottom_components
           ' %info t,_501,Welding_Seams, 焊缝,Welding_Seams
           ' %info t,_502,SPot_Welding,点焊,Spot_Welding
           ' %info t,_503,Adhesive,胶水,adhesive
           ' %info C,_G600,Grou_fasteners,紧固件组合,Group_Fastener
           ' %info C,_G700,others,其他组件,others
           
      ' %info c,_Abandon,Abandoned,废案,Abandoned
      ' %info c,_Patterns,Fasteners,紧固件阵列,Fasteners_Pattern
Private prj
Sub CATMain()
    prj = KCL.GetInput("请输入项目名称"): If prj = "" Then Exit Sub
    Dim Tree As Object: Set Tree = ParsePn(getDecCode())
    Dim PStack As Object: Set PStack = KCL.InitDic
    Dim k, oprd As Object, ref As Object, fast As Object
    
    For Each k In Tree.keys
        Set oprd = AddNode(PStack, Tree(k))
        If InStr(1, k, "_ref", 1) > 0 Then Set ref = oprd
        If InStr(1, k, "_Patterns", 1) > 0 Then Set fast = oprd
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
    Dim oprd, par, TP$: TP = "Product"
    If L < 1 Then L = 1
    If L = 1 Then
        Set oprd = CATIA.Documents.Add("Product").Product:  Set PStack(1) = oprd
    Else
        Set par = PStack(IIf(PStack.Exists(L - 1), L - 1, 1))
        If D.Exists("Type") Then
            If VBA.UCase(VBA.Trim(D("Type"))) = "T" Then TP = "Part"
            If VBA.UCase(VBA.Trim(D("Type"))) = "C" Then TP = "Component"
        End If
        If TP = "Component" Then Set oprd = par.Products.AddNewProduct("") Else Set oprd = par.Products.AddNewComponent(TP, "")
       Set PStack(L) = oprd
    End If
    
    On Error Resume Next
    oprd.Name = D("Name")
    With oprd.ReferenceProduct
        .PartNumber = prj & D("PartNumber"): .Nomenclature = D("Nomenclature"): .Definition = D("Definition")
    End With
    oprd.Update
    Set AddNode = oprd
End Function
Function getDecCode()
    On Error Resume Next
    Dim m As Object: Set m = KCL.GetApc().ExecutingProject.VBProject.VBE.Activecodepane.codemodule
    If Not m Is Nothing Then If m.CountOfDeclarationLines > 0 Then getDecCode = m.Lines(1, m.CountOfDeclarationLines)
End Function

Private Function ParsePn(C$) As Object
    Dim RE As Object, m, lst, curL%, H(20) As Integer, curI%
    Set RE = CreateObject("VBScript.RegExp"): Set lst = KCL.InitDic(1)
    RE.Global = True: RE.MultiLine = True: RE.Pattern = "^(\s*)'\s*%info\s+([^,]*),+([^,]*),+([^,]*),+([^,]*),+([^,\r\n]*).*$"
    If RE.test(C) Then
        H(0) = -1: H(1) = 0
        For Each m In RE.Execute(C)
            curI = Len(m.SubMatches(0))
            curL = GetLev(curI, curL, H)
            Dim D: Set D = KCL.InitDic(1)
            D.Add "Level", curL: D.Add "Type", VBA.Trim(m.SubMatches(1)): D.Add "PartNumber", VBA.Trim(m.SubMatches(2))
            D.Add "Nomenclature", VBA.Trim(m.SubMatches(3)): D.Add "Definition", VBA.Trim(m.SubMatches(4)): D.Add "Name", VBA.Trim(m.SubMatches(5))
            lst.Add D("PartNumber"), D
        Next
    End If
    Set ParsePn = lst
End Function

Private Function GetLev(ByVal i As Integer, ByVal L As Integer, ByRef H() As Integer) As Integer
    If L = 0 Or i > H(L) Then
        L = L + 1: If L > UBound(H) Then L = UBound(H)
        H(L) = i
    Else
        While L > 1 And H(L) > i: L = L - 1: Wend
    End If
    GetLev = L
End Function



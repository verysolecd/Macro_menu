Attribute VB_Name = "ASM_NomOrder"

'{GP:3}
'{Ep:InsNameorder}
'{Caption:实例名排序}
'{ControlTipText: 选择产品后按层级对产品的实例名排序}
'{BackColor:}
'  %UI Label  lblInfo  若需新实例名请输入 #FFFF00
'  %UI TextBox   txt_oName 实例名
'  %UI Button btn_LV  执行：当前层级
'  %UI Button btn_full  执行：全局所有该件

Option Explicit
Private mWD As Cls_DynaWD
Private mPrd As Object
Private targetPN
Private cnt As Long
Private mName As String
Private Const mdlname As String = "ASM_NomOrder"

Sub InsNameorder()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    cnt = 0
    Set mWD = New Cls_DynaWD
    mWD.ShowToolbar mdlname
End Sub


' 读取工具栏文本框的实时内容：为空或仍为占位提示时返回空串
' (工具栏非模态模式下 Cls_DynaWD 不回填 Results 字典，必须直接读窗体上的实时控件，
'  否则 Results 会触发 EnsureExecuted 弹出冗余模态框，导致"需点击两次"且读不到输入)
'=== 辅助：优先读 CATIA 当前 Selection，空才弹选择器 ===
Private Function getSelPrd() As Object
    Set getSelPrd = Nothing
    Dim sel As Selection: Set sel = CATIA.ActiveDocument.Selection
    If sel.count > 0 Then
        On Error Resume Next
            Set getSelPrd = sel.item(1).LeafProduct
        On Error GoTo 0
    End If
    If getSelPrd Is Nothing Then
        Set getSelPrd = KCL.selPrd("请选择要重命名的产品")
    End If
End Function

Private Sub inttext()
    mName = ""
    If mWD Is Nothing Then Exit Sub
    Dim txt As Object: Set txt = mWD.GetCtrlbyname("txt_oName")
    If txt Is Nothing Then Exit Sub
    Dim s As String: s = Trim$(CStr(txt.text))
    If s <> "" And s <> "实例名" Then mName = s
End Sub

Sub btn_LV_click()
    Dim selPrd As Object: Set selPrd = getSelPrd()
    If selPrd Is Nothing Then Exit Sub
    inttext
    cnt = RenameInstByPN_Level(selPrd, mName)
    showcnt selPrd.partNumber
End Sub

Sub btn_full_click()
    Dim selPrd As Object: Set selPrd = getSelPrd()
    If selPrd Is Nothing Then Exit Sub
    Dim rootPrd As Object: Set rootPrd = CATIA.ActiveDocument.Product
    inttext
    cnt = RenameInstByPN_Tree(rootPrd, selPrd.partNumber, mName)
    showcnt selPrd.partNumber
End Sub


Private Sub showcnt(ByVal targetPN As String)
    If cnt = 0 Then
                MsgBox "未找到零件号 [" & targetPN & "] 的可重命名实例，未执行重命名。", vbInformation
            Else
                MsgBox "重命名完成！零件号 [" & targetPN & "]，共处理 " & cnt & " 个实例", vbInformation
            End If
End Sub


Public Function RenameInstByPN_Level(ByVal oprd As Object, Optional ByVal mName As String = "") As Long
    RenameInstByPN_Level = 0
    If oprd Is Nothing Then Exit Function
    ' 当前层级：优先取选中产品的同级集合，仅重命名与选中产品同零件号的实例；
    ' 若选中为顶层根产品（其父级不是 Products），则取根产品的第一层子实例集合，
    ' 取消零件号过滤（RenameChildrenInstByPN 内部会按各自零件号分组重命名）
    If TypeOf oprd.Parent Is Products Then
        RenameInstByPN_Level = RenameChildrenInstByPN(oprd.Parent, False, oprd.partNumber, mName)
    Else
        Dim children As Object
        On Error Resume Next
        Set children = oprd.Products
        On Error GoTo 0
        If children Is Nothing Then Exit Function
        RenameInstByPN_Level = RenameChildrenInstByPN(children, False, "", mName)
    End If
End Function

Public Function RenameInstByPN_Tree(ByVal rootPrd As Object, ByVal targetPN As String, Optional ByVal mName As String = "") As Long
    RenameInstByPN_Tree = 0
    If rootPrd Is Nothing Or targetPN = "" Then Exit Function
    Dim children As Products
    On Error Resume Next
    Set children = rootPrd.Products
    On Error GoTo 0
    If children Is Nothing Then Exit Function
    If children.count < 1 Then Exit Function
    RenameInstByPN_Tree = RenameChildrenInstByPN(children, False, targetPN, mName)
    Dim i As Long
    For i = 1 To children.count
        RenameInstByPN_Tree = RenameInstByPN_Tree + RenameInstByPN_Tree(children.item(i), targetPN, mName)
    Next i
End Function

Private Function RenameChildrenInstByPN(ByVal children As Object, _
                                        ByVal onlyDup As Boolean, _
                                        ByVal filterPN As String, _
                                        ByVal mName As String) As Long
    RenameChildrenInstByPN = 0
    If children Is Nothing Then Exit Function
    Dim n As Long: n = children.count
    If n < 1 Then Exit Function
    Dim sibs() As Object: ReDim sibs(1 To n)
    Dim i As Long
    On Error Resume Next
    For i = 1 To n
        Set sibs(i) = children.item(i)
        If sibs(i) Is Nothing Then
            RenameChildrenInstByPN = 0
            Err.Clear
            Exit Function
        End If
    Next i
    On Error GoTo 0
    Dim groups As Object: Set groups = KCL.InitDic
    Dim pn As String
    For i = 1 To n
        pn = sibs(i).ReferenceProduct.partNumber
        If Not groups.Exists(pn) Then groups.Add pn, KCL.Initlst
        groups(pn).Add i
    Next i
    Dim used As Object: Set used = KCL.InitDic
    For i = 1 To n
        If Not used.Exists(sibs(i).Name) Then used.Add sibs(i).Name, True
    Next i
    Dim cnt As Long: cnt = 0
    Dim gkey As Variant, lst As Object, j As Long, base As String
    For Each gkey In groups.keys
        If filterPN <> "" And gkey <> filterPN Then GoTo NextGroup
        Set lst = groups(gkey)
        If onlyDup And lst.count < 2 Then GoTo NextGroup

        For j = 0 To lst.count - 1
            If used.Exists(sibs(lst(j)).Name) Then used.Remove sibs(lst(j)).Name
        Next j
        If mName <> "" Then
            base = mName
        Else
            base = Trim$(sibs(lst(0)).ReferenceProduct.Nomenclature)
            If base = "" Then base = Trim$(sibs(lst(0)).ReferenceProduct.partNumber)
        End If
        Dim k As Long, cand As String
        k = 0
        For j = 0 To lst.count - 1
            k = k + 1
            cand = NextUniqueInstName(base, k, used)
            On Error Resume Next
            sibs(lst(j)).Name = cand
            If Err.Number = 0 Then
                If Not used.Exists(cand) Then used.Add cand, True
                cnt = cnt + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        Next j
NextGroup:
    Next gkey

    RenameChildrenInstByPN = cnt
End Function
Private Function NextUniqueInstName(ByVal base As String, ByVal startK As Long, ByVal used As Object) As String
    Dim k As Long: k = startK
    Dim cand As String
    Do
        cand = base & "." & Format$(k, "00")
        If Not used.Exists(cand) Then Exit Do
        k = k + 1
    Loop
    NextUniqueInstName = cand
End Function









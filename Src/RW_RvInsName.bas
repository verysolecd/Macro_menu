Attribute VB_Name = "RW_RvInsName"
'{GP:1}
'{Ep:InsRename}
'{Caption:实例名按零件号重排序}
'{ControlTipText:选择一个产品取其零件号，将同级同零件号实例重命名为 nomenclature.01..n，各层级独立}
'{BackColor:}
'UI 在代码中动态构建(AddUIElement)，第一行显示当前选择的零件号，顺序即自上而下的布局顺序

Option Explicit
Private mWD As Cls_DynaWD
Private Const mdlname As String = "RW_RvInsName"

Sub InsRename()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    Dim oprd As Object
    Set oprd = KCL.SelPrd("请选择要重命名实例名的产品")
    If oprd Is Nothing Then Exit Sub

    Dim targetPN As String: targetPN = oprd.ReferenceProduct.partNumber
    If targetPN = "" Then
        MsgBox "所选产品零件号为空，无法处理。", vbExclamation
        Set oprd = Nothing
        Exit Sub
    End If

    Set mWD = New Cls_DynaWD
    mWD.Title = "实例名按零件号重排序"
    mWD.AddUIElement "Label", "lblPN", "当前选择零件号：" & targetPN
    mWD.AddUIElement "Label", "lblPrefix", "实例名前缀（留空则按原 Nomenclature 命名，将自动追加 .01..n）："
    mWD.AddUIElement "TextBox", "txt_Prefix", ""
    mWD.AddUIElement "Label", "lblInfo", "请勾选重命名范围："
    mWD.AddUIElement "CheckBox", "chk_Level", "仅当前父层级(同级)"
    mWD.AddUIElement "CheckBox", "chk_Tree", "所有出现该零件号的层级(每层独立)"
    mWD.AddUIElement "Button", "btnOK", "确定"
    mWD.AddUIElement "Button", "btncancel", "取消"
    mWD.Show

    Select Case mWD.btnClicked
        Case "btnOK"
            Dim doLevel As Boolean, doTree As Boolean
            doLevel = mWD.Results("chk_Level")
            doTree = mWD.Results("chk_Tree")
            If Not doLevel And Not doTree Then
                MsgBox "请至少勾选一个重命名范围！", vbExclamation
                Set oprd = Nothing
                Exit Sub
            End If
            Dim userPrefix As String
            If IsEmpty(mWD.Results("txt_Prefix")) Then
                userPrefix = ""
            Else
                userPrefix = Trim$(CStr(mWD.Results("txt_Prefix")))
            End If
            Dim cnt As Long: cnt = 0
            Dim rootPrd As Object: Set rootPrd = CATIA.ActiveDocument.Product
            If doTree Then
                cnt = RenameInstByPN_Tree(rootPrd, targetPN, userPrefix)
            ElseIf doLevel Then
                cnt = RenameInstByPN_Level(oprd, userPrefix)
            End If
            If cnt = 0 Then
                MsgBox "未找到零件号 [" & targetPN & "] 的可重命名实例，未执行重命名。", vbInformation
            Else
                MsgBox "重命名完成！零件号 [" & targetPN & "]，共处理 " & cnt & " 个实例（" & _
                       IIf(doTree, "各层级独立编号", "仅当前父层级") & "）。", vbInformation
            End If
        Case Else
    End Select
    Set oprd = Nothing
End Sub

Public Function RenameInstByPN_Level(ByVal oprd As Object, Optional ByVal userPrefix As String = "") As Long
    RenameInstByPN_Level = 0
    If oprd Is Nothing Then Exit Function
    If Not TypeOf oprd.Parent Is Products Then Exit Function
    Dim targetPN As String: targetPN = oprd.ReferenceProduct.partNumber
    If targetPN = "" Then Exit Function
    RenameInstByPN_Level = RenameChildrenInstByPN(oprd.Parent, False, targetPN, userPrefix)
End Function

Public Function RenameInstByPN_Tree(ByVal rootPrd As Object, ByVal targetPN As String, Optional ByVal userPrefix As String = "") As Long
    RenameInstByPN_Tree = 0
    If rootPrd Is Nothing Or targetPN = "" Then Exit Function
    Dim children As Products
    On Error Resume Next
    Set children = rootPrd.Products
    On Error GoTo 0
    If children Is Nothing Then Exit Function
    If children.count < 1 Then Exit Function

    RenameInstByPN_Tree = RenameChildrenInstByPN(children, False, targetPN, userPrefix)

    Dim i As Long
    For i = 1 To children.count
        RenameInstByPN_Tree = RenameInstByPN_Tree + RenameInstByPN_Tree(children.item(i), targetPN, userPrefix)
    Next i
End Function

Private Function RenameChildrenInstByPN(ByVal children As Object, _
                                        ByVal onlyDup As Boolean, _
                                        ByVal filterPN As String, _
                                        ByVal userPrefix As String) As Long
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

        If userPrefix <> "" Then
            base = userPrefix
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


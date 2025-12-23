Attribute VB_Name = "OTH_NewBH"

'Attribute VB_Name = "OTH_NewBH"
'{GP:6}
'{Ep:CATMain}
'{Caption:新电池箱体}
'{ControlTipText:新建一个电池箱体的结构树}
'{BackColor:}
'======零件号信息

' %info Product,_Prj_Housing_Asm,Project_HousingAsm,箱体组件,HousingAsm
' %info Product,_Pack,Pack_system,整包方案,Pack_system
' %info Product,_Packaging,packaging,包络定义,packaging
' %info Product,_000,Upper_Housing_Asm,上箱体总成,Upper_Housing_Asm
' %info Part,_001,Upper_Housing,上箱体,Upper_Housing
' %info Product,_1000,Lower_Housing_Asm,下箱体总成,Lower_Housing_Asm
' %info Part,_ref,Ref,参考,Ref
' %info Part,_1100,Frames,框架组件,Frames
' %info Part,_1200,Brkts,支架类组件,Brkts
' %info Part,_1300,Cooling_system,液冷组件,Cooling_system
' %info Part,_1400,Bottom_components,底部组件,Bottom_components
' %info Part,_2001,Welding_Seams,焊缝,Welding_Seams
' %info Part,_2002,SPot_Welding,点焊,Spot_Welding
' %info Part,_2003,Adhesive,胶水,adhesive
' %info Product,_4000,Grou_fasteners,紧固件组合,Group_Fastener
' %info Part,_5000,others,其他组件,others
' %info Product,_Abandon,Abandoned,废案,Abandoned
' %info Product,_Patterns,Fasteners,紧固件阵列,Fasteners_Pattern


Private prj
Sub CATMain()
    Dim oPrd As Object, Tree As Object
    Dim ref As Object, fast As Object
    Dim imsg As String, kArray, i As Integer
    Dim Tdict As Object
    
    imsg = "请输入项目名称": prj = KCL.GetInput(imsg)
    If prj = "" Then Exit Sub
    
    Set Tree = initTree()
    kArray = Tree.keys
    
    ' 定义父级堆栈
    Dim ParentStack As Object
    Set ParentStack = KCL.InitDic(1)
    
    '==== 动态层级创建逻辑 ===
    For i = 0 To Tree.count - 1
        Set Tdict = Tree(kArray(i))
        
        Set oPrd = AddNodeToTree(ParentStack, Tdict)
        
        ' 捕获特殊对象用于后续操作
        Dim keyName As String
        keyName = CStr(kArray(i))
        If InStr(1, keyName, "_ref", vbTextCompare) > 0 Then Set ref = oPrd
        If InStr(1, keyName, "_Patterns", vbTextCompare) > 0 Then Set fast = oPrd
        
        Set oPrd = Nothing
    Next
    
    '=== 后续处理：Ref 复制到 Fastener ===
    If (Not ref Is Nothing) And (Not fast Is Nothing) Then
        Dim osel
        Set osel = CATIA.ActiveDocument.Selection: osel.Clear
        osel.Add ref: osel.Copy
        osel.Clear
        
        Dim otp
        Set otp = CATIA.ActiveDocument.Selection: otp.Clear
        otp.Add fast: otp.Paste
        otp.Clear
    End If
    
    Set allPN = KCL.InitDic(vbTextCompare): allPN.RemoveAll
    ' 注意：这里如果 rootprd 没有在 CATMain 里显式定义， recurInitPrd 可能找不到对象
    ' 由于 AddNodeToTree 第一个创建的就是 Root，我们可以从 ParentStack(1) 获取
    If ParentStack.Exists(1) Then
        Call recurInitPrd(ParentStack(1))
    End If
End Sub

' --- 新增：处理节点创建和堆栈管理的独立函数 ---
Function AddNodeToTree(ByRef ParentStack As Object, ByVal Tdict As Object) As Object
    Dim curLevel As Integer
    Dim oPrd As Object
    Dim parentPrd As Object
    Dim oDoc As Document
    
    ' 1. 获取层级
    If Tdict.Exists("Level") Then
        curLevel = CInt(Tdict("Level"))
    Else
        curLevel = 1
    End If
    If curLevel < 1 Then curLevel = 1
    
    ' 2. 创建节点逻辑
    If curLevel = 1 Then
        ' Level 1: 创建根产品
        Set oDoc = CATIA.Documents.Add("Product")
        Set oPrd = oDoc.Product
        
        ' 将根节点存入堆栈 Level 1
        ParentStack.Add 1, oPrd
    Else
        ' Level > 1: 从堆栈找爸爸
        If ParentStack.Exists(curLevel - 1) Then
            Set parentPrd = ParentStack(curLevel - 1)
        Else
            ' 回退保护: 找不到父级就挂在 Level 1
            If ParentStack.Exists(1) Then
                Set parentPrd = ParentStack(1)
            Else
                ' 极端情况：连 Level 1 都没有 (不应该发生)
                MsgBox "Error: Root product not found for " & Tdict("PartNumber")
                Exit Function
            End If
        End If
        
        ' 决定类型
        Dim compType As String
        compType = "Product" ' 默认
        If Tdict.Exists("Type") Then
            If UCase(Trim(Tdict("Type"))) = "PART" Then
                compType = "Part"
            End If
        End If
        
        Set oPrd = parentPrd.Products.AddNewComponent(compType, "")
         
        ' 更新堆栈
        If ParentStack.Exists(curLevel) Then
            Set ParentStack(curLevel) = oPrd
        Else
            ParentStack.Add curLevel, oPrd
        End If
    End If
    
    ' 3. 设置属性
    Call newPn(oPrd, Tdict)
    
    Set AddNodeToTree = oPrd
End Function

Function initTree()
    DecCode = getDecCode()
    Set initTree = ParsePn(DecCode)
End Function

Private Function ParsePn(ByVal code As String) As Object
    Dim regex, matches, match, lst As Object
    Dim compInfo As Object 
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        ' 匹配格式: [缩进] ' %info <Type>,<Pn>,<Nomenclature>,<Definition>,<Name>
        ' Group 1: (\s*) 捕获行首缩进，用于计算层级
        ' Group 2: ([^,]+) 捕获 Type
        ' ... 后续字段
        .Pattern = "^(\s*)'\s*%info\s+([^,]+),+([^,]+),+([^,]+),+([^,]+),+([^,\r\n]+).*$"
    End With
    If regex.Test(code) Then
        Set matches = regex.Execute(code)
        Set lst = KCL.InitDic(1)
        
        Dim i As Integer
        Dim curIndent As Integer
        Dim curLevel As Integer
        
        ' IndentHistory: 索引是 Level, 值是该 Level 对应的缩进长度
        ' Level 0 预留为 -1 (比任何缩进都小)
        Dim IndentHistory(20) As Integer
        IndentHistory(0) = -1
        IndentHistory(1) = 0 ' 默认 Level 1 缩进为 0 (或者第一行的缩进)
        
        curLevel = 0 
        
        For Each match In matches
            Dim matchIndentIdx
            matchIndentIdx = 0 ' 正则 group 0 是缩进 (\s*)
            
            curIndent = Len(match.SubMatches(matchIndentIdx))
            
            If curLevel = 0 Then
                ' 第一行，强制定义为 Level 1
                curLevel = 1
                IndentHistory(1) = curIndent
            Else
                If curIndent > IndentHistory(curLevel) Then
                    ' 缩进增加 -> 进入下一层
                    curLevel = curLevel + 1
                    IndentHistory(curLevel) = curIndent
                ElseIf curIndent = IndentHistory(curLevel) Then
                    ' 缩进相同 -> 同级
                    ' curLevel 不变
                Else
                    ' 缩进减少 ->回退查找之前是哪一层
                    Dim j As Integer
                    Dim found As Boolean
                    found = False
                    For j = curLevel - 1 To 1 Step -1
                        If curIndent >= IndentHistory(j) Then 
                            If curIndent = IndentHistory(j) Then
                                curLevel = j
                                found = True
                                Exit For
                            End If
                        End If
                    Next
                    
                    If Not found Then
                        For j = curLevel - 1 To 1 Step -1
                             If IndentHistory(j) <= curIndent Then
                                curLevel = j
                                Exit For
                             End If
                        Next
                    End If
                End If
            End If
            
            Set compInfo = KCL.InitDic
            compInfo.Add "Level", curLevel
            ' 移除了 Sequence 字段
            compInfo.Add "Type", Trim(match.SubMatches(1))
            compInfo.Add "PartNumber", Trim(match.SubMatches(2))
            compInfo.Add "Nomenclature", Trim(match.SubMatches(3))
            compInfo.Add "Definition", Trim(match.SubMatches(4))
            compInfo.Add "Name", Trim(match.SubMatches(5))
            
            lst.Add compInfo("PartNumber"), compInfo
        Next
    End If
    Set ParsePn = lst
End Function
Sub newPn(oPrd, Dic)
    Call KCL.showdict(Dic)
    
    ' 1. 先修改 Instance Name (实例名)
    ' 对于第三层深度的组件，先改这个通常更稳健
    On Error Resume Next
        oPrd.Name = Dic("Name")
    On Error GoTo 0
    
    ' 2. 再获取 ReferenceProduct 修改零件号等属性
    ' 修改 Reference 属性不会影响 Instance 对象的有效性，但反过来有时会
    Dim refPrd
    Set refPrd = oPrd.ReferenceProduct
    
    Dim k
    k = "PartNumber"
    refPrd.PartNumber = prj & Dic(k)
    
    k = "Nomenclature"
    refPrd.Nomenclature = Dic(k)
    
    k = "Definition"
    refPrd.Definition = Dic(k)
    
    oPrd.Update
End Sub
Function getDecCode()
    Dim DecCnt, DecCode
     Dim Apc As Object: Set Apc = KCL.GetApc()
         Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
         On Error Resume Next
          Dim mdl: Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
             Error.Clear
         On Error GoTo 0
    If mdl Is Nothing Then Exit Function
    DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then Exit Function
    DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
    getDecCode = DecCode
End Function

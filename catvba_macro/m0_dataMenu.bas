Attribute VB_Name = "m0_dataMenu"
'Attribute VB_Name = "Cat_Macro_Menu_Model"
' 此代码用于获取宏菜单所需的配置信息并展示菜单界面
Const FormTitle = "Macro"
'----- 菜单的配置信息 ---------------------------------------
' 菜单的显示类型
' True - 非模态显示  False - 模态显示
Private Const MENU_SHOW_TYPE = True
' 菜单的隐藏类型
' True - 隐藏菜单按钮  False - 显示菜单按钮
Private Const MENU_HIDE_TYPE = False
' 菜单分组的配置信息
' 请根据需要修改
'{ 分组编号 : 分组标题 }
' 示例配置
Private Const groupName = _
            "{1 : 图纸处理 }" & _
            "{2 : 零件建模 }" & _
            "{3 : 总成装配 }" & _
            "{4 : 读取修改 }" & _
            "{5 : BOM处理}"
'-----------------------------------------------------------------
Option Explicit
'----- 配置参数 请勿修改除非必要 -----------------------
' 菜单分组映射表
Private PageMap As Object
' 标签映射表
Private TagMap As Object                    ' 分组编号标签
Private Const TAG_S = "{"                   ' 配置开始标签
Private Const TAG_D = ":"                   ' 配置分隔标签
Private Const TAG_E = "}"                   ' 配置结束标签
Private Const TAG_GROUP = "gp"              ' 分组编号标签
Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_ENTRY_DEF = "CATMain"     ' 入口点默认值
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
'-----------------------------------------------------------------
' 菜单入口点
Sub CATMain()
    Set PageMap = Get_KeyValue(groupName, True)
    Dim ButtonInfos As Object
    Set ButtonInfos = Get_ButtonInfo()
    If ButtonInfos Is Nothing Then
        MsgBox "未找到可用的宏信息", vbExclamation + vbOKOnly
        Exit Sub
    End If
    ' 对按钮信息进行排序
    Dim SoLst As Object
    Set SoLst = To_SortedList(ButtonInfos)
    If SoLst Is Nothing Then Exit Sub
    ' 显示菜单界面
    Dim Menu
    Set Menu = New Cat_Macro_Menu_View
    Call Menu.Set_FormInfo(SoLst, PageMap, FormTitle, MENU_HIDE_TYPE)
    
    If MENU_SHOW_TYPE Then
        Menu.Show vbModeless
    Else
        Menu.Show vbModal
    End If
End Sub
'******* 辅助函数 *********
' 获取宏按钮的配置信息
' 参数  :
' 返回值: lst(Dict)
Private Function Get_ButtonInfo() As Object
    Set Get_ButtonInfo = Nothing
    
    Dim Apc As Object: Set Apc = GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    Dim PjtPath As String: PjtPath = ExecPjt.DisplayName
    
    Dim AllComps As Object
    Set AllComps = GetModuleLst(ExecPjt.ProjectItems.VBComponents)
    If AllComps Is Nothing Then Exit Function
    
    Dim Comp As Object 'VBComponent
    Dim Mdl As Object 'CodeModule
    Dim DecCode As String
    Dim DecCnt As Long
    Dim MdlInfo As Object
    Dim CanExecMethod As String
    Dim BtnInfos As Object: Set BtnInfos = KCL.InitLst()
    
    For Each Comp In AllComps
        Set Mdl = Comp.CodeModule
        ' 获取声明行数
        DecCnt = Mdl.CountOfDeclarationLines
        If DecCnt < 1 Then GoTo continue
        ' 获取声明代码
        DecCode = Mdl.Lines(1, Mdl.CountOfDeclarationLines)
        ' 获取配置信息
        Set MdlInfo = Get_KeyValue(DecCode)
        If MdlInfo Is Nothing Then GoTo continue
        ' 检查分组信息
        If Not MdlInfo.Exists(TAG_GROUP) Then GoTo continue
        If IsNumeric(MdlInfo(TAG_GROUP)) Then
            MdlInfo(TAG_GROUP) = CLng(MdlInfo(TAG_GROUP))
        Else
            GoTo continue
        End If
        Debug.Print TypeName(MdlInfo(TAG_GROUP)) & " : " & MdlInfo(TAG_GROUP)
        If Not PageMap.Exists(MdlInfo(TAG_GROUP)) Then GoTo continue
        
        ' 检查入口点方法
        CanExecMethod = vbNullString
        If MdlInfo.Exists(TAG_ENTRYPNT) Then
            If Exist_Method(Mdl, MdlInfo(TAG_ENTRYPNT)) Then
                CanExecMethod = MdlInfo(TAG_ENTRYPNT)
            Else
                GoTo Try_TAG_ENTRY_DEF
            End If
        Else
Try_TAG_ENTRY_DEF:
            If Exist_Method(Mdl, TAG_ENTRY_DEF) Then
                 CanExecMethod = TAG_ENTRY_DEF
            End If
        End If
        If CanExecMethod = vbNullString Then GoTo continue
        Set MdlInfo = Push_Dic(MdlInfo, TAG_ENTRYPNT, CanExecMethod)
        Set MdlInfo = Push_Dic(MdlInfo, TAG_PJTPATH, PjtPath)
        Set MdlInfo = Push_Dic(MdlInfo, TAG_MDLNAME, Mdl.Name)
        BtnInfos.Add MdlInfo
continue:
    Next
    If BtnInfos.Count < 1 Then Exit Function
    Set Get_ButtonInfo = BtnInfos
End Function
' 向字典中添加或更新键值对
' 参数  : Dict,vri,vri
' 返回值: Dict
Private Function Push_Dic(ByVal Dic As Object, _
                          ByVal Key As Variant, _
                          ByVal item As Variant) As Object
    If Dic.Exists(Key) Then
        Dic(Key) = item
    Else
        Dic.Add Key, item
    End If
    Set Push_Dic = Dic
End Function
' 从字符串中提取配置信息 - 键转换为长整型
' 参数  : str,Opt_bool
' 返回值: Dict
Private Function Get_KeyValue( _
                    ByVal txt As String, _
                    Optional ByVal KeyToLong As Boolean = False) _
                    As Object
    Set Get_KeyValue = Nothing
    Dim Reg As Object
    Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Pattern = TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E
        .Global = True
    End With
    
    Dim Matches As Object
    Set Matches = Reg.Execute(txt)
    Set Reg = Nothing
    
    If Matches.Count < 1 Then Exit Function
    
    Dim Dic As Object: Set Dic = KCL.InitDic(vbTextCompare)
    Dim Match As Object, SubMatchs As Object
    Dim Key As Variant, Var As Variant
    
    For Each Match In Matches
        Set SubMatchs = Match.SubMatches
        
        If SubMatchs.Count < 2 Then GoTo continue
        
        Key = Trim(Replace(SubMatchs(0), """", ""))
        If Len(Key) < 1 Then GoTo continue
        If KeyToLong Then Key = CLng(Key)
        
        Var = Trim(Replace(SubMatchs(1), """", ""))
        If Len(Var) < 1 Then GoTo continue
        
        Set Dic = Push_Dic(Dic, Key, Var)
continue:
    Next
    
    If Dic.Count < 1 Then Exit Function
    
    Set Get_KeyValue = Dic
End Function
' 将按钮信息按分组排序
' 参数  :lst(Dict)
' 返回值: Dict(lst(Dict))
Private Function To_SortedList(ByVal Infos As Object) As Object
    Set To_SortedList = Nothing
    
    Dim SoLst As Object
    Set SoLst = CreateObject("System.Collections.SortedList")
    Dim lst As Object
    
    Dim Info As Object
    For Each Info In Infos
        If SoLst.ContainsKey(Info(TAG_GROUP)) = True Then
            SoLst(Info(TAG_GROUP)).Add Info
        Else
            Set lst = KCL.InitLst()
            lst.Add Info
            SoLst.Add Info(TAG_GROUP), lst
        End If
    Next
    
    If SoLst.Count < 1 Then Exit Function
    
    ' 按模块名称排序
    Dim i As Long
    Dim InfoDic As Object: Set InfoDic = KCL.InitDic(vbTextCompare)
    For i = 0 To SoLst.Count - 1
        InfoDic.Add SoLst.GetKey(i), Sort_by(SoLst.GetByIndex(i))
    Next
    
    Set To_SortedList = InfoDic
End Function
' 按模块名称排序
' 参数  :lst(Dict)
' 返回值: lst(Dict)
Private Function Sort_by(ByVal lst As Object) As Object
    Dim tmp As Object
    Dim i As Long, j As Long
    Set tmp = lst(0)
    For i = 0 To lst.Count - 1
        For j = lst.Count - 1 To i Step -1
            If lst(i)(TAG_MDLNAME) > lst(j)(TAG_MDLNAME) Then
                Set tmp = lst(i)
                Set lst(i) = lst(j)
                Set lst(j) = tmp
            End If
        Next j
    Next i
    Set Sort_by = lst
End Function
'******* APC/VBE *********
' 获取APC对象
' 参数  :
' 返回值: obj-IApc
Private Function GetApc() As Object
    Set GetApc = Nothing
    
    ' 获取VBA版本对应的COM对象名称
    Dim COMObjectName$
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        MsgBox "不支持当前VBA版本", vbExclamation + vbOKOnly
        Exit Function
    #End If
    
    ' 获取APC对象
    Dim Apc As Object: Set Apc = Nothing
    On Error Resume Next
        Set Apc = CreateObject(COMObjectName)
    On Error GoTo 0
    
    If Apc Is Nothing Then
        MsgBox "无法获取MSAPC.Apc对象", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    Set GetApc = Apc
End Function
' 检查代码模块中是否存在指定方法 - 不检查私有方法
' 参数  : obj-CodeModule,str
' 返回值: Boolean
Private Function Exist_Method(ByVal CodeMdl As Object, _
                              ByVal Name As String) As Boolean
    Dim tmp As Long
    On Error Resume Next
        tmp = CodeMdl.ProcBodyLine(Name, 0)
    On Error GoTo 0
    Exist_Method = tmp > 0
    Err.Number = 0
End Function
' 获取标准模块列表
' 参数  : obj-VBComponents
' 返回值: lst(obj-VBComponent)
' vbext_ComponentType
' 1-vbext_ct_StdModule 2-vbext_ct_ClassModule 3-vbext_ct_MSForm
Private Function GetModuleLst(ByVal Itms As Object) As Object
    Set GetModuleLst = Nothing
    Dim lst As Object: Set lst = KCL.InitLst()
    Dim Itm As Object
    For Each Itm In Itms
        If Not Itm.Type = 1 Then GoTo continue 'vbext_ComponentType
        lst.Add Itm
continue:
    Next
    If lst.Count < 1 Then Exit Function
    Set GetModuleLst = lst
End Function


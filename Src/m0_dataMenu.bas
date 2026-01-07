Attribute VB_Name = "m0_dataMenu"
'Attribute VB_Name = "Cat_Macro_Menu_Model"
' 此代码用于获取宏菜单所需的配置信息并展示菜单界面
Const formTitle = "键盘造车手"
'----- 菜单的配置信息 ---------------------------------------
' True - 非模态显示  False - 模态显示
Private Const Menu_Modeless = True
' True - 隐藏菜单按钮  False - 显示菜单按钮
Private Const MENU_HIDE_TYPE = True
' 菜单分组的配置信息
' 请根据需要修改
'{ 分组编号 : 分组标题 }
' 示例配置
Private Const groupName = _
            "{1 : R&W }" & _
            "{2 : BOM }" & _
            "{3 : ASM }" & _
            "{4 : MDL }" & _
            "{5 : DRW }" & _
            "{6 : OTRS}"
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
    Set PageMap = Get_KeyValue(groupName, True)  '获取page编号和名称对应map  ：1  R&W 2...
   
        showdict PageMap
    
    Dim ButtonInfos As Object
    Set ButtonInfos = Get_ButtonInfo() '获取所有的具有可执行按钮的模块信息dic
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
    Call Menu.Set_FormInfo(SoLst, PageMap, formTitle, MENU_HIDE_TYPE)
    If Menu_Modeless Then
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
    Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    ' ==   获取catvba文件path
    Dim PjtPath As String: PjtPath = ExecPjt.DisplayName
    ' ==  获取所有Module，这里还未识别是否是菜单module
    Dim AllComps As Object
    Set AllComps = GetModuleLst(ExecPjt.ProjectItems.VBComponents)
    If AllComps Is Nothing Then Exit Function
    
    Dim comp As Object 'VBComponent
    Dim mdl As Object 'CodeModule
    Dim DecCode As String
    Dim DecCnt As Long
    Dim MdlInfo As Object
    Dim CanExecMethod As String
    
    Dim BtnInfos As Object: Set BtnInfos = KCL.InitLst()
    
    For Each comp In AllComps
        Set mdl = comp.codemodule
        DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then GoTo Continue
        DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
        Set MdlInfo = Get_KeyValue(DecCode) ' 将声明与page pair对比，获取配置信息
        If MdlInfo Is Nothing Then GoTo Continue
        If Not MdlInfo.Exists(TAG_GROUP) Then GoTo Continue ' 检查分组信息
        If IsNumeric(MdlInfo(TAG_GROUP)) Then   '这里 MdlInfo(TAG_GROUP)是字典检查taggroup对应的值
            MdlInfo(TAG_GROUP) = CLng(MdlInfo(TAG_GROUP))
        Else
            GoTo Continue
        End If
'            Debug.Print TypeName(MdlInfo(TAG_GROUP)) & " : " & MdlInfo(TAG_GROUP)
        If Not PageMap.Exists(MdlInfo(TAG_GROUP)) Then GoTo Continue '若mpages编号不包含该分组值则下一个
        ' 检查入口点方法
        CanExecMethod = vbNullString
        If MdlInfo.Exists(TAG_ENTRYPNT) Then  '若module声明包含EP即函数入口则
            If Exist_Method(mdl, MdlInfo(TAG_ENTRYPNT)) Then   '，检查是否有入口函数，MdlInfo(TAG_ENTRYPNT)是获取对应ep的函数名字
                CanExecMethod = MdlInfo(TAG_ENTRYPNT) '获取可执行函数名
            Else
                GoTo Try_TAG_ENTRY_DEF
            End If
        Else
        
Try_TAG_ENTRY_DEF:
            If Exist_Method(mdl, TAG_ENTRY_DEF) Then
                 CanExecMethod = TAG_ENTRY_DEF
            End If
            
        End If
        
        If CanExecMethod = vbNullString Then GoTo Continue
        Set MdlInfo = Push_Dic(MdlInfo, TAG_ENTRYPNT, CanExecMethod)
        Set MdlInfo = Push_Dic(MdlInfo, TAG_PJTPATH, PjtPath) '字典存储项目路径
        Set MdlInfo = Push_Dic(MdlInfo, TAG_MDLNAME, mdl.Name) '字典存储模块名称
        
        '一个具有可执行按钮的模块的例子
        ' MlInfo
        ' Compar elModeTextCompare
        ' Count
        ' Item 1 "GP"
        ' Item 2 "EP"
        ' Item 3 “Caption"
        ' Item 4 "ControlTipText”
        ' Item 5 “背景颜色
        ' Item 6 "pjt_path"
        ' Item 7 "mdl_name”
        
        BtnInfos.Add MdlInfo
    Debug.Print showdict(MdlInfo)
Continue:
    Next
    If BtnInfos.count < 1 Then Exit Function
    Set Get_ButtonInfo = BtnInfos  '获取所有的具有可执行按钮的模块信息dic
    
End Function
' 向字典中添加或更新键值对
' 参数  : Dict,vri,vri
' 返回值: Dict
Private Function Push_Dic(ByVal Dic As Object, _
                          ByVal key As Variant, _
                          ByVal item As Variant) As Object
    If Dic.Exists(key) Then
        Dic(key) = item
    Else
        Dic.Add key, item
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
    
    Dim matches As Object
    Set matches = Reg.Execute(txt)
    Set Reg = Nothing
    
    If matches.count < 1 Then Exit Function
    
    Dim Dic As Object: Set Dic = KCL.InitDic(vbTextCompare)
    Dim match As Object, SubMatchs As Object
    Dim key As Variant, Var As Variant
    
    For Each match In matches
        Set SubMatchs = match.SubMatches
        If SubMatchs.count < 2 Then GoTo Continue
        ' ==  获取编号
        key = VBA.Trim(VBA.Replace(SubMatchs(0), """", "")) 'trim 取消前后空格， replace 删除中间空格
        If Len(key) < 1 Then GoTo Continue  '若key为空进入下一个循环
        If KeyToLong Then key = CLng(key)  'Clng转换为long类型
            ' ==  获取编号对应page
        Var = VBA.Trim(VBA.Replace(SubMatchs(1), """", ""))  'trim 取消前后空格， replace 删除中间空格
        If Len(Var) < 1 Then GoTo Continue
        Set Dic = Push_Dic(Dic, key, Var)
Continue:
    Next
    
    If Dic.count < 1 Then Exit Function
    
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
    Dim info As Object
    For Each info In Infos
        If SoLst.ContainsKey(info(TAG_GROUP)) = True Then
            SoLst(info(TAG_GROUP)).Add info
        Else
            Set lst = KCL.InitLst()
            lst.Add info
            SoLst.Add info(TAG_GROUP), lst
        End If
    Next
    
    If SoLst.count < 1 Then Exit Function
    
    ' 按模块名称排序
    Dim i As Long
    Dim InfoDic As Object: Set InfoDic = KCL.InitDic(vbTextCompare)
    For i = 0 To SoLst.count - 1
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
    For i = 0 To lst.count - 1
        For j = lst.count - 1 To i Step -1
            If lst(i)(TAG_MDLNAME) > lst(j)(TAG_MDLNAME) Then
                Set tmp = lst(i)
                Set lst(i) = lst(j)
                Set lst(j) = tmp
            End If
        Next j
    Next i
    Set Sort_by = lst
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
    Dim itm As Object
    For Each itm In Itms
        If Not itm.Type = 1 Then GoTo Continue 'vbext_ComponentType
        lst.Add itm
Continue:
    Next
    If lst.count < 1 Then Exit Function
    Set GetModuleLst = lst
End Function



    
 



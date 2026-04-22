Attribute VB_Name = "A0TEST_Engine"
'===========================================================================
' A0TEST_Engine - Cls_DynaUIEngine 验证测试模块
'
' 使用方法:
'   1. 将此文件和 Cls_DynaUIEngine.cls 导入 CATIA VBA 工程
'   2. 在 VBA 编辑器中逐个运行以下 Sub 进行验证:
'       Test1_ModalPopup   - 验证模态弹窗 (等效 OTH_Minibox 的UI交互)
'       Test2_Toolbar      - 验证非模态工具栏 (等效 OTH_ivhideshow)
'       Test3_Alert        - 验证信息弹窗
'       Test4_CodeDriven   - 验证纯代码构建UI (新增能力)
'       Test5_MainMenu     - 验证主菜单数据收集 (LoadFromMenuTags)
'===========================================================================
Option Explicit

'===========================================================================
' 测试1: 模态弹窗 (对比 OTH_Minibox 的使用方式)
'
' 老代码:
'   Dim oFrm: Set oFrm = KCL.newFrm(mdlname)
'   oFrm.Show
'   Select Case oFrm.BtnClicked
'
' 新代码:
'   Dim oEng: Set oEng = KCL.newEngine(mdlname)
'   oEng.Show
'   Select Case oEng.ClickedButton
'===========================================================================
'{GP:99}
'{Ep:Test1_ModalPopup}
'{Caption:测试-模态弹窗}
'{ControlTipText:验证Cls_DynaUIEngine模态模式}

' %UI Label lbl_info  [Engine测试] 模态弹窗
' %UI CheckBox chk_opt1  选项A
' %UI CheckBox chk_opt2  选项B
' %UI TextBox  txt_input  请输入测试文本
' %UI Button btnOK     确定
' %UI Button btnCancel 取消

Private Const mdlname As String = "A0TEST_Engine"

Sub Test1_ModalPopup()
    ' ===== 使用新引擎 =====
    Dim oEng As Cls_DynaUIEngine
    Set oEng = KCL.newEngine(mdlname)
    oEng.Show
    
    ' 检查结果
    If oEng.IsCancelled Then
        MsgBox "用户取消了操作", vbInformation"
    Else
        Dim msg As String
        msg = "点击的按钮: " & oEng.ClickedButton & vbCrLf
        
        ' 遍历所有结果
        Dim res As Object: Set res = oEng.Results
        Dim k As Variant
        For Each k In res.keys
            msg = msg & k & " = " & res(k) & vbCrLf
        Next
        MsgBox msg, vbInformation, "模态弹窗结果""
    End If
    
    Set oEng = Nothing
End Sub

'===========================================================================
' 测试2: 非模态工具栏 (对比 OTH_ivhideshow 的使用方式)
'
' 老代码:
'   Set g_frm = KCL.newFrm(mdlname)
'   g_frm.ShowToolbar mdlname, mapMdl, mapFunc
'
' 新代码:
'   Dim oEng: Set oEng = KCL.newEngine()
'   oEng.ShowToolbar mdlname, mapMdl, mapFunc
'===========================================================================
Sub Test2_Toolbar()
    ' 构建按钮->模块映射 和 按钮->函数映射
    Dim mapMdl As Object: Set mapMdl = KCL.InitDic
    Dim mapFunc As Object: Set mapFunc = KCL.InitDic
    
    ' 将每个按钮的 Name 映射到本模块和对应的 _Click 函数
    mapMdl("btnOK") = mdlname
    mapFunc("btnOK") = "btnOK_toolbar_click"
    mapMdl("btnCancel") = mdlname
    mapFunc("btnCancel") = "btnCancel_toolbar_click"
    
    ' 使用新引擎
    Dim oEng As New Cls_DynaUIEngine
    oEng.ShowToolbar mdlname, mapMdl, mapFunc
    ' 工具栏是非模态的，代码会立即返回到这里
    ' 按钮点击时会通过 ExecuteScript 调用下面的 _Click 函数
End Sub

' 工具栏按钮的回调函数
Sub btnOK_toolbar_click()
    MsgBox "工具栏 [确定] 被点击!", vbInformation
End Sub
Sub btnCancel_toolbar_click()
    MsgBox "工具栏 [取消] 被点击!", vbInformation
End Sub

'===========================================================================
' 测试3: Alert 信息弹窗 (对比 cls_dynaFrm.Alert)
'
' 老代码:
'   oFrm.Alert "100.00 x 200.00 x 50.00"
'
' 新代码:
'   oEng.Alert "100.00 x 200.00 x 50.00"
'===========================================================================
Sub Test3_Alert()
    ' 先执行模态弹窗获取结果，再用 Alert 展示
    Dim oEng As Cls_DynaUIEngine
    Set oEng = KCL.newEngine(mdlname)
    
    ' Alert 可以直接调用，无需先 LoadFromModuleName
    ' 但因为 Alert 内部需要一个 TextBox 来显示内容，
    ' 我们先加载模块定义（含 txt_input TextBox）
    oEng.Alert "测试包络尺寸: 123.45 x 678.90 x 42.00" & vbCrLf & _
               "点击 [复制并关闭] 按钮可复制到剪贴板""
    
    Set oEng = Nothing
End Sub

'===========================================================================
' 测试4: 纯代码构建UI (全新能力 —— 无需任何 %UI 注释)
'
' 这是老架构完全做不到的，新引擎独有的扩展方式
'===========================================================================
Sub Test4_CodeDriven()
    Dim oEng As New Cls_DynaUIEngine
    
    ' 手动设置标题
    oEng.Title = "动态构建的界面"
    
    ' 手动追加控件 —— 完全不依赖任何模块注释
    oEng.AddUIElement "Label", "lbl_header", "请选择导出格式:"
    oEng.AddUIElement "CheckBox", "chk_stp", "STEP (.stp)"
    oEng.AddUIElement "CheckBox", "chk_igs", "IGES (.igs)"
    oEng.AddUIElement "TextBox", "txt_path", "C:\output"
    oEng.AddUIElement "Button", "btnExport", "开始导出", "#336699"
    oEng.AddUIElement "Button", "btnCancel", "取消"
    
    ' 一行显示
    oEng.Show
    
    ' 收结果
    If Not oEng.IsCancelled Then
        MsgBox "导出按钮: " & oEng.ClickedButton & vbCrLf & _
               "STP: " & oEng.Results("chk_stp") & vbCrLf & _
               "IGES: " & oEng.Results("chk_igs") & vbCrLf & _
               "路径: " & oEng.Results("txt_path"), _
               vbInformation, "纯代码UI结果"
    End If
    
    Set oEng = Nothing
End Sub

'===========================================================================
' 测试5: 主菜单数据收集 (LoadFromMenuTags)
'
' 这个测试验证新引擎能否正确扫描工程中所有带 {GP:} 标签的模块
' 并生成与 Cat_Macro_Menu_View.Set_FormInfo 兼容的数据结构
'
' 老代码 (A00_Menu):
'   Set MenuItems = GetMenuItems()         ' 扫描
'   Set SoLst = OrganizeForView(MenuItems) ' 分组排序
'
' 新代码:
'   Dim oEng As New Cls_DynaUIEngine
'   Set SoLst = oEng.LoadFromMenuTags(PageMap)  ' 一步到位
'===========================================================================
Sub Test5_MainMenu()
    ' 1. 解析 GroupName 定义 (复用 A00_Menu 中的格式)
    Dim GroupDef As String
    GroupDef = "{1 : R&W }" & _
               "{3 : ASM }" & _
               "{4 : MDL }" & _
               "{5 : DRW }" & _
               "{7: CATIA }" & _
               "{6 : OTRS}"
    
    Dim PageMap As Object
    Set PageMap = ParseGroupDef(GroupDef)
    
    ' 2. 用新引擎收集所有菜单项
    Dim oEng As New Cls_DynaUIEngine
    Dim SoLst As Object
    Set SoLst = oEng.LoadFromMenuTags(PageMap)
    
    If SoLst Is Nothing Then
        MsgBox "未扫描到任何有效菜单项", vbExclamation"
        Exit Sub
    End If
    
    ' 3. 打印扫描结果 (验证数据正确性)
    Dim msg As String
    Dim grpKey As Variant, items As Object, item As Object
    For Each grpKey In SoLst.keys
        msg = msg & "=== 组 " & grpKey
        If PageMap.Exists(grpKey) Then msg = msg & " (" & PageMap(grpKey) & ")"
        msg = msg & " ===" & vbCrLf
        
        Set items = SoLst(grpKey)
        Dim i As Long
        For i = 0 To items.count - 1
            Set item = items(i)
            msg = msg & "  " & item("mdl_name") & " -> " & item("ep") & vbCrLf
        Next
    Next
    
    MsgBox msg, vbInformation, "LoadFromMenuTags 扫描结果"
    
    ' 4. [可选] 实际渲染到主菜单窗体 —— 取消注释以下代码验证
    '    Dim Menu As Cat_Macro_Menu_View
    '    Set Menu = New Cat_Macro_Menu_View
    '    Call Menu.Set_FormInfo(SoLst, PageMap, "键盘造车手(Engine)", True)
    '    Menu.Show vbModeless
    
    Set oEng = Nothing
End Sub

' --- 辅助: 解析 GroupName 字符串 (从 A00_Menu.get_Tagcfg 复制) ---
Private Function ParseGroupDef(ByVal txt As String) As Object
    Dim dic As Object: Set dic = KCL.InitDic(vbTextCompare)
    Dim Reg As Object: Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Pattern = "{(.*?):(.*?)}"
        .Global = True
    End With
    Dim matches As Object: Set matches = Reg.Execute(txt)
    Dim match As Object
    Dim KEY As Variant, val As Variant
    For Each match In matches
        If match.SubMatches.count >= 2 Then
            KEY = Trim(match.SubMatches(0))
            val = Trim(match.SubMatches(1))
            If IsNumeric(KEY) Then KEY = CLng(KEY)
            If dic.Exists(KEY) Then dic(KEY) = val Else dic.Add KEY, val
        End If
    Next
    Set ParseGroupDef = dic
End Function

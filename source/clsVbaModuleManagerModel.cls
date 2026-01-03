VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVbaModuleManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DEBUG_MODE = False                        '调试模式
Private Const APP_NAME = "ModuleBridge"                 '应用名称
Private Const VERSION = "0.0.6"                         '版本
Private Const THIS_PROJECT_NAME = "CAT_MENU" '"ModuleBridgeProject" '此项目名
Private Const USER_JSON_NAME = "user_data.json"         '数据保存文件名
Private Const SEPARATOR = "|"                           '路径分隔符-VBA-JSON对策
Private Const SOURCE_FOLDER_NAME = "source"             '在CATVBA内创建时的文件夹
Private mVbe As Object                                  'VBA编辑器
Private mUtil As clsVBAUtilityLib                       '工具类
Private mProjectLst As collection                       '项目列表
Private mUserDict As Object                             '用户数据字典
Private mJson As clsJsonConverter                       'Json解析器
Private Sub Class_Initialize()
    Set mUtil = New clsVBAUtilityLib
    mUtil.DebugMode = DEBUG_MODE
    Set mVbe = mUtil.GetVbe()
    Set mJson = New clsJsonConverter
    Set mProjectLst = get_project_list()
    Set mUserDict = get_user_data()
End Sub
Private Sub Class_Terminate()
    Dim path As String
    path = get_this_dir_path() & "\" & USER_JSON_NAME
    Call cleanUp_userdata
    save_json path, mUserDict
End Sub
'***属性***
'应用标题
Property Get title() As String
    title = APP_NAME & " Ver" & VERSION
End Property
'此项目名
Property Get project_name() As String
   project_name = THIS_PROJECT_NAME
End Property
'***方法***
'获取项目
'return:VBProject集合
Function get_project_list() As collection
    Dim lst As collection
    Set lst = New collection
    Dim proj As Object
    For Each proj In mVbe.vBProjects
        lst.Add proj
    Next
    Set get_project_list = lst
End Function
'获取项目名
'return:项目名
Function get_project_name_list() As collection
    Dim lst As collection
    Set lst = New collection
    Dim proj As Object
    For Each proj In mProjectLst
        lst.Add proj.Name
    Next
    Set get_project_name_list = lst
End Function
'获取项目内的模块
'param: projectIdx-项目索引
'return: 模块集合
Function get_module_list( _
        ByVal projectIdx As Long) As collection
    Dim lst As collection
    Set lst = New collection
    Set get_module_list = lst
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    Dim comp As Object
    For Each comp In proj.VBComponents
        If is_target_module_type(comp) Then
            lst.Add comp
        End If
    Next
End Function
'获取项目内的模块名
'param: projectIdx-项目索引
'return: 模块名集合
Function get_module_name_list( _
        ByVal projectIdx As Long) As collection
    Dim lst As collection
    Set lst = New collection
    Set get_module_name_list = lst
    Dim comp As Object
    For Each comp In get_module_list(projectIdx)
        If is_target_module_type(comp) Then
            lst.Add comp.Name
        End If
    Next
End Function
'获取指定模块内函数名列表
'param: comp-模块
'return: 函数名字符串
Private Function get_fanction_names( _
        ByVal comp As Object) As String
    Dim lst As collection
    Set lst = New collection
    Dim buf As String
    Dim i As Long
    With comp.codemodule
        For i = 1 To .CountOfLines
            If buf <> .ProcOfLine(i, 0) Then
                buf = .ProcOfLine(i, 0)
                lst.Add "  [" + buf + "]"
            End If
        Next i
    End With
    get_fanction_names = Join( _
        mUtil.CollectionToArrayLet(lst), _
        vbCrLf _
    )
End Function
'是否为vba模块
'vbext_ct_StdModule 1 标准模块
'vbext_ct_ClassModule 2 类模块
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX设计器
'vbext_ct_Document 100 文档模块
Private Function is_target_module_type( _
        ByVal module As Object) As Boolean
    Select Case module.Type
        Case 1, 2, 3
            is_target_module_type = True
        Case Else
            is_target_module_type = False
    End Select
End Function
'获取模块信息
'param: projectIdx-项目索引
'return: 项目的概要信息
Function get_module_info( _
        ByVal projectIdx As Long, _
        ByVal moduleName As String) As String
    get_module_info = vbNullString
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    Dim comp As Object
    On Error Resume Next
        Set comp = proj.VBComponents(moduleName)
    On Error GoTo 0
    If comp Is Nothing Then Exit Function
    '项目总行数
    Dim totalCharCount As Long
    totalCharCount = 0
    Dim compX As Object
    For Each compX In proj.VBComponents
        totalCharCount = totalCharCount + GetValidCharacterCount(compX)
    Next
    get_module_info = _
        "项目名称:" & proj.Name & vbCrLf & _
        "总体步数:" & totalCharCount & vbCrLf & _
        "文件路径:" & proj.Filename & vbCrLf & _
        "模块名称:" & moduleName & "." & get_extension(comp) & vbCrLf & _
        "模块步数:" & GetValidCharacterCount(comp) & vbCrLf & _
        "函数列表:" & vbCrLf & get_fanction_names(comp)
End Function
'获取组件除去注释部分的字符数
'param: vbComp-目标组件
Private Function GetValidCharacterCount( _
        ByVal VBComp As Object) As Long
    Dim CodeMod As Object ' CodeModule
    Dim totalChars As Long
    Dim i As Long
    Dim codeLine As String
    ' 获取CodeModule对象
    Set CodeMod = VBComp.codemodule
    totalChars = 0
    ' 检查每行并排除注释
    For i = 1 To CodeMod.CountOfLines
        codeLine = VBA.Trim(CodeMod.Lines(i, 1))
        If Left(codeLine, 1) <> "'" Then
            totalChars = totalChars + Len(codeLine)
        End If
    Next i
    ' 返回结果
    GetValidCharacterCount = totalChars
End Function
'删除项目内的vba模块并导入
'param: projectIdx-项目索引
Sub import_project( _
        ByVal projectIdx As Long)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    Dim key As String
    key = Replace(proj.Filename, "\", SEPARATOR)
    If Not mUserDict.Exists(key) Then Exit Sub
    Dim path As String
    path = mUserDict(key)("path")
    If Not mUtil.GetFso().FolderExists(path) Then Exit Sub
    Call remove_modules(projectIdx)
    Call import_modules(proj, path)
End Sub
'根据项目名获取文件夹-如果没有则创建
'param: projectIdx-项目索引
'param: parentDirPath-父文件夹路径
'return: 创建的文件夹路径
Function get_dir_by_project_name( _
        ByVal projectIdx As Long, _
        ByVal parentDirPath As String) As String
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    Dim path As String
    path = parentDirPath & "\" & proj.Name
    Dim fso As Object
    Set fso = mUtil.GetFso()
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    get_dir_by_project_name = path
End Function
'导入文件夹内的文件
'param: proj-项目
'param: dirPath-导入的文件夹路径
Private Sub import_modules( _
        ByVal proj As Object, _
        ByVal dirPath As String)
    Dim fileLst As collection
    Set fileLst = get_target_file_path_list( _
        dirPath, _
        Array( _
            "cls", _
            "frm", _
            "bas" _
        ) _
    )
    Dim path As Variant ' String
    For Each path In fileLst
        On Error Resume Next
            proj.VBComponents.import path
        On Error GoTo 0
    Next
End Sub
'获取指定路径内指定扩展名的文件
'param: path-文件夹路径
'param: targetExtensions-目标扩展名
'return: 文件路径集合
Private Function get_target_file_path_list( _
        ByVal path As String, _
        ByVal targetExtensions As Variant) As collection
    Dim lst As collection
    Set lst = New collection
    Set get_target_file_path_list = lst
    Dim fso As Object
    Set fso = mUtil.GetFso()
    Dim dir As Object
    Set dir = fso.GetFolder(path)
    Dim file As Object
    Dim ext As String
    For Each file In dir.Files
        ext = fso.GetExtensionName(file.Name)
        If Not UBound(filter(targetExtensions, ext)) < 0 Then
            lst.Add file.path
        End If
    Next
End Function
'删除项目内的vba模块
'param: projectIdx-项目索引
Private Sub remove_modules( _
        ByVal projectIdx As Long)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    Dim lst As collection
    Set lst = get_module_name_list(projectIdx)
    Dim Name As Variant ' String
    Dim comp As Object
    For Each Name In get_module_name_list(projectIdx)
        Set comp = proj.VBComponents(Name)
        proj.VBComponents.Remove comp
    Next
End Sub
'导出项目内的vba模块
'param: projectIdx-项目索引
'param: dirPath-文件夹路径
'param: isOpen(opt)-是否打开?
Sub export_project( _
        ByVal projectIdx As Long, _
        ByVal dirPath As String, _
        Optional ByVal isOpen As Boolean = True)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    
    Dim Name 'As String
    Dim comp As Object
    For Each Name In get_module_name_list(projectIdx)
        Set comp = proj.VBComponents(Name)
        On Error Resume Next
        comp.Export dirPath & "\" & comp.Name & "." & get_extension(comp)
        If Not Err.Number = 0 Then
            MsgBox "导出失败，请删除.git以外的文件后重试"
            Exit Sub
        End If
        On Error GoTo 0
    Next
    update_userdata proj, dirPath
    If isOpen Then
        open_folder_dy_explorer dirPath
    End If
End Sub
'导出项目内的vba模块到CATVBA文件夹内
'param: projectIdx-项目索引
'param: isOpen(opt)-是否打开?
Sub export_project_child_folder( _
        ByVal projectIdx As Long, _
        Optional ByVal isOpen As Boolean = True)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    Dim fso As Object
    Set fso = mUtil.GetFso()
    Dim dirPath As String
    dirPath = fso.GetParentFolderName(proj.Filename) & "\" & SOURCE_FOLDER_NAME
    If Not fso.FolderExists(dirPath) Then
        fso.CreateFolder (dirPath)
    End If
    Call export_project(projectIdx, dirPath)
End Sub
'覆盖导出项目内的vba模块
'param: projectIdx-项目索引
Sub overwriting_project( _
        ByVal projectIdx As Long)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    Dim key As String
    key = Replace(proj.Filename, "\", SEPARATOR)
    Dim dirPath As String
    dirPath = mUserDict(key)("path")
    Dim fso As Object
    Set fso = mUtil.GetFso()
    '找不到的话删除就可以了吗？
    If Not fso.FolderExists(dirPath) Then Exit Sub
    Call remove_module_files(dirPath)
    Call export_project(projectIdx, dirPath)
End Sub
'仅打开项目文件夹
'param: projectIdx-项目索引
Sub open_folder( _
        ByVal projectIdx As Long)
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Sub
    Dim key As String
    key = Replace(proj.Filename, "\", SEPARATOR)
    Dim dirPath As String
    dirPath = mUserDict(key)("path")
    Dim fso As Object
    Set fso = mUtil.GetFso()
    '找不到的话删除就可以了吗？
    If Not fso.FolderExists(dirPath) Then Exit Sub
    Call open_folder_dy_explorer(dirPath)
End Sub
'更新用户数据
'由于vba-json对"\\"的处理不好，所以替换为SEPARATOR
'param: proj-项目
'param: path-保存目录路径
Private Sub update_userdata( _
        ByVal proj As Object, _
        ByVal path As String)
    Dim key As String
    key = Replace(proj.Filename, "\", SEPARATOR)
    If mUserDict.Exists(key) Then
        mUserDict.Remove key
    End If
    Dim info As Object
    Set info = mUtil.InitDict()
    info.Add "name", proj.Name
    info.Add "path", path
    mUserDict.Add key, info
    Call cleanUp_userdata
End Sub
'获取导出用的扩展名
'https://learn.microsoft.com/ja-jp/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
'vbext_ct_StdModule 1 标准模块
'vbext_ct_ClassModule 2 类模块
'vbext_ct_MSForm 3 Microsoft Forms
'vbext_ct_ActiveXDesigner 11 ActiveX设计器
'vbext_ct_Document 100 文档模块
Private Function get_extension( _
        ByVal module As Object) As String
    Dim extension As String
    Select Case module.Type
        Case 1 'vbext_ct_StdModule
            extension = "bas"
        Case 2 'vbext_ct_ClassModule
            extension = "cls"
        Case 3 'vbext_ct_MSForm
            extension = "frm"
        Case Else
            extension = ""
    End Select
    get_extension = extension
End Function
'删除模块文件
Private Sub remove_module_files( _
        ByVal dirPath As String)
    Dim fso As Object
    Set fso = mUtil.GetFso()
    If Not fso.FolderExists(dirPath) Then Exit Sub
    Dim Paths As collection
    Set Paths = get_target_file_path_list( _
        dirPath, _
        Array( _
            "cls", _
            "frm", _
            "frx", _
            "bas" _
        ) _
    )
    Dim path As Variant ' String
    For Each path In Paths
        On Error Resume Next
            fso.DeleteFile path
        On Error GoTo 0
    Next
    '等待
    Call mUtil.Wait(1)
End Sub
'获取活动项目
'return: 此项目对象
Function get_this_project() As Object
    Set get_this_project = Nothing
    Dim thisProj As Object
    Dim proj As Object
    For Each proj In mVbe.vBProjects
        If proj.Name = THIS_PROJECT_NAME Then
            Set get_this_project = proj
            Exit Function
        End If
    Next
End Function
'获取活动项目文件夹路径
'return: 此项目的文件夹路径
Private Function get_this_dir_path() As String
    get_this_dir_path = mUtil.GetFso().GetParentFolderName( _
        get_this_project().Filename _
    )
End Function
'通过资源管理器打开文件夹
'param: dirPath-文件夹路径
Private Sub open_folder_dy_explorer( _
        ByVal dirPath As String)
    On Error Resume Next
        shell "C:\Windows\Explorer.exe " & dirPath, vbNormalFocus
    On Error GoTo 0
End Sub
'获取用户数据
'return: dict
Private Function get_user_data() As Object
    Dim path As String
    path = get_this_dir_path() & "\" & USER_JSON_NAME
    mUtil.Dump "userdata_path:" & path
    Dim dict As Object
    If mUtil.GetFso().FileExists(path) Then
        Set dict = load_json(path)
        dict("ver") = VERSION
    Else
        Set dict = mUtil.InitDict()
        dict.Add "ver", VERSION
    End If
    Set get_user_data = dict
End Function
'清理用户数据
Private Sub cleanUp_userdata()
    Dim fso As Object
    Set fso = mUtil.GetFso()
    Dim removeLst As collection
    Set removeLst = New collection
    Dim filePath As String
    Dim dirPath As String
    Dim key As Variant 'String
    For Each key In mUserDict.keys()
        If key = "ver" Then GoTo Continue
        filePath = Replace(key, SEPARATOR, "\")
        If Not fso.FileExists(filePath) Then
            removeLst.Add key
            GoTo Continue
        End If
        dirPath = mUserDict(key)("path")
        If Not fso.FolderExists(dirPath) Then
            removeLst.Add key
            GoTo Continue
        End If
Continue:
    Next
    For Each key In removeLst
        mUserDict.Remove key
    Next
End Sub
'项目是否曾被保存过
'param: projectIdx-项目索引
Function has_user_data( _
        ByVal projectIdx As Long) As Boolean
    has_user_data = False
    Dim proj As Object
    On Error Resume Next
        Set proj = mProjectLst.item(projectIdx)
    On Error GoTo 0
    If proj Is Nothing Then Exit Function
    has_user_data = mUserDict.Exists(Replace(proj.Filename, "\", SEPARATOR))
End Function
'写入json
'param: path-json文件路径
'param: 写入的dict
Private Sub save_json( _
        ByVal path As String, _
        ByVal dict As Object)
    Call mUtil.WriteFile( _
        path, _
        mJson.ConvertToJson(dict, 3) _
    )
End Sub
'读取json
'param: path-json文件路径
'return: dict
Private Function load_json( _
        ByVal path As String) As Object
    Set load_json = Nothing
    If Not mUtil.GetFso().FileExists(path) Then Exit Function
    Dim dict As Object
    Set dict = mJson.ParseJson(mUtil.ReadFile(path))
    Set load_json = dict
End Function


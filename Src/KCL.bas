Attribute VB_Name = "KCL"

Option Explicit


Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr

#If VBA7 Then
    ' 现有的声明
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If
' 常量定义
Private Const SW_MAXIMIZE = 3
Private Const SW_NORMAL = 1
Private mSW& ' 秒表开始时间

Public Type Bomline
    level           As Integer        ' 层级
    partNumber      As String    ' 件号
    Nomenclature    As String  ' 英文名称
    Definition      As String    ' 中文名称
    InstanceName    As String  ' 实例名
    Quantity        As Long        ' 数量
    Mass            As Double          ' 单重
    Material        As String      ' 材质
    Thickness       As Double     ' 厚度
    Density         As Double       ' 密度
    UserProp1       As String
    UserProp2       As String
End Type

Public Type ParamItem
    Name            As String
    ParamType       As String
    Value           As Variant
    target          As Object    ' 指向 CATIA Parameter 对象
    Description     As String
End Type

Public rootDoc
Public rootPrd  As Object
Public xlAPP As Object
Public gwb As Object
Public gws  As Object
Public pdm As New Cls_PDM
Public xlm As New Cls_XLM
Public g_allPN As Object
Public g_Picpath
Public g_Btn

'*****计时相关函数*****
' 启动秒表
Private Const mdlname As String = "KCL"
Sub SW_Start()
    mSW = timeGetTime
End Sub
' 获取计时时间
''' @return:Double(Unit:s)
Function SW_GetTime#()
    SW_GetTime = IIf(mSW = 0, -1, (timeGetTime - mSW) * 0.001)
End Function

'*****CATIA相关函数*****=================================================
' 循环选择项目
Sub LoopSel()
    Dim msg$: msg = "请选择项目 : 按ESC键退出"
    Dim SI As AnyObject
    Dim doc As Document: Set doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(msg)
        If IsNothing(SI) Then Exit Do
        Stop
    Loop
End Sub
' 检查是否可以执行操作
''' @param:DocTypes-array(string),string 指定可执行操作的文档类型
''' @return:Boolean
Function CanExecute(ByVal docTypes As Variant) As Boolean
    CanExecute = False
    If CATIA.Windows.count < 1 Then: MsgBox "没有打开的窗口": Exit Function
    If VarType(docTypes) = vbString Then
      docTypes = VBA.LCase(docTypes)
        docTypes = Split(docTypes, ",") '过滤器转数组
    End If
    If Not checkFilterType(docTypes) Then Exit Function '过滤器检查，非数组则退出
    
   Dim ErrMsg As String: ErrMsg = "不支持当前文档类型。" + vbNewLine + "(仅支持" + Join(docTypes, ",") + ")"
    CanExecute = checkDocType(docTypes)
    If Not CanExecute Then MsgBox ErrMsg, vbExclamation + vbOKOnly
End Function
' 检查文档类型
''' @param:DocTypes-指定的文档类型
Function checkDocType(ByVal docTypes As Variant)
    If CATIA.Windows.count < 1 Then: MsgBox "没有打开的窗口": Exit Function
    checkDocType = False
    If VarType(docTypes) = vbString Then
         docTypes = VBA.LCase(docTypes)
        docTypes = VBA.Split(docTypes, ",") '过滤器转数组
    End If
    If Not checkFilterType(docTypes) Then Exit Function '过滤器检查，非数组则退出
    
    Dim ActDoc As Document
        On Error Resume Next
            Set ActDoc = CATIA.ActiveDocument
        On Error GoTo 0
    If ActDoc Is Nothing Then: Exit Function
     If UBound(filter(docTypes, VBA.LCase(TypeName(ActDoc)))) < 0 Then '此处filter函数是VBA中比较后返回数组的函数
        Exit Function
    End If
    checkDocType = True
End Function
Function SelPrd(ByVal msg$, _
                    Optional ByVal filter As Variant = Empty)
    If checkDocType("partdocument") Then
        Set SelPrd = CATIA.ActiveDocument.Product
        Exit Function
    End If
    Dim se As SelectedElement
    Set se = SelectElement(msg, filter)
    If IsNothing(se) Then
        Set SelPrd = se
    Else
        Set SelPrd = se.LeafProduct
    End If
End Function
' 选择项目 /产品/零件/body/几何图形集等
''' @param:Msg-提示信息
''' @param:Filter-array(string),string 选择过滤器(默认为AnyObject)
''' @return:AnyObject
Function SelectItem(ByVal msg$, _
                    Optional ByVal filter As Variant = Empty)
    Dim se As SelectedElement
    Set se = SelectElement(msg, filter)
    If IsNothing(se) Then
        Set SelectItem = se
    Else
        Set SelectItem = se.Value
    End If
End Function
' 选择元素
''' @param:Msg-提示信息
''' @param:Filter-array(string),string 选择过滤器(默认为AnyObject)
''' @return:SelectedElement
Function SelectElement(ByVal msg$, _
                           Optional ByVal filter As Variant = Empty) ' _
                           As SelectedElement
    If IsEmpty(filter) Then filter = Array("AnyObject")
    If VarType(filter) = vbString Then filter = strToAry(filter)
    If Not checkFilterType(filter) Then Exit Function
    Dim sel: Set sel = CATIA.ActiveDocument.Selection: sel.Clear
    Select Case sel.SelectElement2(filter, msg, False)
        Case "Cancel", "Undo", "Redo"
             Set SelectElement = Nothing
            Exit Function
    End Select
    Set SelectElement = sel.item(1)
    sel.Clear
End Function

Function Selectmulti(ByVal msg$, _
                           Optional ByVal filter As Variant = Empty) ' _
                           As SelectedElement
    If IsEmpty(filter) Then filter = Array("AnyObject")
    If VarType(filter) = vbString Then filter = strToAry(filter)
    If Not checkFilterType(filter) Then Exit Function
    Dim sel: Set sel = CATIA.ActiveDocument.Selection: sel.Clear
    Select Case sel.SelectElement3(filter, msg, True, 2, True)
        Case "Cancel", "Undo", "Redo"
             Set Selectmulti = Nothing
            Exit Function
    End Select
    Set Selectmulti = sel
End Function

Public Function defPrd()
    Dim res As Integer
    Dim tmp
    Set tmp = Nothing
    res = MsgBox("“是”选择，“否”根节点，“取消”退出", vbYesNoCancel + vbDefaultButton2, "选择产品")
    Select Case res
        Case 6: Set tmp = KCL.SelPrd("请选择目标产品")
        Case 7: Set tmp = CATIA.ActiveDocument.Product
        Case Else: Set tmp = Nothing
    End Select
    Set defPrd = tmp
End Function

' 获取内部名称
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(aoj)
    If IsNothing(aoj) Then
        GetInternalName = Empty: Exit Function
    End If
    GetInternalName = aoj.GetItem("ModelElement").internalName
End Function

' 获取obj的特定类型的父对象
''' @param:anyObj-AnyObject
''' @param:T-String
''' @return:AnyObject
Function GetParent_Of_T( _
                        ByVal anyObj As AnyObject, _
                        ByVal t As String) As AnyObject
    Dim anyObjName, parentName As String
    On Error Resume Next
        Set anyObj = asDisp(anyObj)
        anyObjName = anyObj.Name
        parentName = anyObj.Parent.Name
    On Error GoTo 0
    If TypeName(anyObj) = TypeName(anyObj.Parent) And _
       anyObjName = parentName Then
        Set GetParent_Of_T = Nothing
        Exit Function
    End If
    If TypeName(anyObj) = t Then
        Set GetParent_Of_T = anyObj
    Else
        Set GetParent_Of_T = GetParent_Of_T(anyObj.Parent, t)
    End If
End Function

Private Function asDisp(o As INFITF.CATBaseDispatch) As INFITF.CATBaseDispatch
    Set asDisp = o
End Function

Public Function get_aPartDoc()
    Set get_aPartDoc = Nothing
    Dim oprt, doc
    For Each doc In CATIA.Documents
        If TypeName(doc) = "PartDocument" Then
            Set get_aPartDoc = doc
            Exit Function
            End If
    Next
End Function

Public Function get_workPartDoc()  'in Assembly
    Set get_workPartDoc = Nothing
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Dim sDocType As String: sDocType = LCase(TypeName(oDoc))
    If sDocType = "partdocument" Then
        Set get_workPartDoc = oDoc
        Exit Function
    ElseIf sDocType <> "productdocument" Then
        Exit Function
    End If
    Dim msel: Set msel = oDoc.Selection
    ' 1. Cache Existing Selection
    Dim bRestore As Boolean: bRestore = False
    Dim cachedSel() As Variant
    Dim i As Integer
    If msel.count > 0 Then
        bRestore = True
        ReDim cachedSel(msel.count - 1)
        For i = 1 To msel.count
           Set cachedSel(i - 1) = msel.item(i).Value
        Next
    End If
    ' 2. Perform Search for Active Context
    msel.Clear
    msel.Search "CATprtSearch.PartFeature,in"
    Dim itemp
    Set itemp = Nothing
    On Error Resume Next
        If msel.count > 0 Then Set itemp = msel.item(1).LeafProduct.ReferenceProduct.Parent '.part
    Err.Clear
    On Error GoTo 0
    
    ' 3. Restore Selection
    msel.Clear
    If bRestore Then
        On Error Resume Next
        For i = 0 To UBound(cachedSel)
            If Not cachedSel(i) Is Nothing Then
                msel.Add cachedSel(i)
            End If
        Next
        On Error GoTo 0
    End If
    Set get_workPartDoc = itemp
End Function


Public Sub toMP()
    On Error Resume Next
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "https://mp.weixin.qq.com/s?__biz=MzU5MTk1MDUwNg==&mid=2247484525&idx=1&sn=554a37aff4bc876424043a9aa5968d6d&scene=21&poc_token=HCUyg2ijuHYXMx810A5yID4tAYIemJFdJ7FpVvew"
    Set shell = Nothing
    If Err.Number <> 0 Then
        MsgBox "无法公众号链接" & vbCrLf & "错误: " & Err.Description, vbExclamation, "链接错误"
    End If
    On Error GoTo 0
End Sub
'========================数组处理==========================================================================
' 获取数组指定范围的元素
''' @param:Ary-Variant(Of Array)
''' @param:StartIdx-Long
''' @param:EndIdx-Long
''' @return:Variant(Of Array)
Function GetRangeAry(ByVal ary As Variant, ByVal startIdx&, ByVal endIdx&) As Variant
    If Not IsArray(ary) Then Exit Function
    If endIdx - startIdx < 0 Then Exit Function
    If startIdx < 0 Then Exit Function
    If endIdx > UBound(ary) Then Exit Function
    Dim rngAry() As Variant: ReDim rngAry(endIdx - startIdx)
    Dim i&
    For i = startIdx To endIdx
        rngAry(i - startIdx) = ary(i)
    Next
    GetRangeAry = rngAry
End Function
' 检查是否为字符串数组
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    IsStringAry = True
End Function
' 将字符串转换为数组变量
Private Function strToAry(ByVal S$) As Variant
    Dim ary As Variant: ary = Split(S, ",")
    Dim oAry() As Variant: ReDim oAry(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        oAry(i) = ary(i)
    Next
    strToAry = oAry
End Function
' 检查过滤器类型是否有效（是字符串数组）
Private Function checkFilterType(ByVal ary As Variant) As Boolean
    checkFilterType = False
    Dim ErrMsg$: ErrMsg = "过滤器类型无效" + vbNewLine + _
                          "需要为Variant(String)类型的数组" + vbNewLine + _
                          "(具体请参考文档)"
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    checkFilterType = True
End Function



'=========================================通用函数========================================
' 检查对象是否为Nothing
''' @param:OJ-Variant(Of Object)
''' @return:Boolean
Function IsNothing(ByVal oj As Variant) As Boolean
    IsNothing = oj Is Nothing
End Function
' 检查对象是否为指定类型
''' @param:OJ-Object
''' @param:T-String
''' @return:Boolean
Function IsObj_T(ByVal oj As Object, ByVal t$) As Boolean
    IsObj_T = IIf(TypeName(oj) = t, True, False)
'    MsgBox TypeName(oj)
End Function
Public Function getItm(iName, colls)
 Dim itm ' 正确声明数组
    Set itm = Nothing
    On Error Resume Next
        Set itm = colls.item(iName)
            Err.Clear
            Err.Number = 0
    On Error GoTo 0
   Set getItm = itm
    Set itm = Nothing
End Function

Public Function SelectQuery(iQuery As String, Optional ByVal iRange = Nothing)
    Set SelectQuery = Nothing
  Dim msel: Set msel = CATIA.ActiveDocument.Selection
  msel.Clear
    If VarType(iQuery) = vbString Then
         iQuery = VBA.LCase(iQuery)
        iQuery = VBA.Split(iQuery, ",") '过滤器转数组
    End If
    If Not checkFilterType(iQuery) Then Exit Function '过滤器检查，非数组则退出
    Dim i As Long, Qry As String
    For i = LBound(iQuery) To UBound(iQuery)
        Qry = iQuery(i)
        If iRange Is Nothing Then
          msel.Search Qry & ",all"
        Else
           msel.Add iRange
           msel.Search Qry & ",sel"
        End If
    Next i
    Set SelectQuery = msel
End Function

Public Function getSearch(ByRef ifilter As Variant, ByRef iDoc)
    Set getSearch = Nothing
    On Error Resume Next
             Dim osel As Selection, i
             Set osel = iDoc.Selection
              osel.Clear
    If LCase(TypeName(ifilter)) = "string" Then
        With osel
            .Clear
            .Search (ifilter)
            .VisProperties.SetShow 1
        End With
    End If
        Set getSearch = osel
End Function

'*****数组相关函数*****
' 合并两个数组
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Variant(Of Array)
Function JoinAry(ByVal ary1 As Variant, ByVal Ary2 As Variant)
    Select Case True
        Case Not IsArray(ary1) And Not IsArray(Ary2)
            JoinAry = Empty: Exit Function
        Case Not IsArray(ary1)
            JoinAry = Ary2: Exit Function
        Case Not IsArray(Ary2)
            JoinAry = ary1: Exit Function
    End Select
    Dim StCount&: StCount = UBound(ary1)
    ReDim Preserve ary1(UBound(ary1) + UBound(Ary2) + 1)
    Dim i&
    If IsObject(Ary2(0)) Then
        For i = StCount + 1 To UBound(ary1)
            Set ary1(i) = Ary2(i - StCount - 1)
        Next
    Else
        For i = StCount + 1 To UBound(ary1)
            ary1(i) = Ary2(i - StCount - 1)
        Next
    End If
    JoinAry = ary1
End Function
' mapping 数组
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function mappedAry(ByVal ary As Variant, ByVal iMap As Variant) As Variant
    If Not IsArray(ary) Or Not IsArray(iMap) Then Exit Function
    Dim CloneAry
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
    Dim ele_Ary()
    Dim mapdata
       mapdata = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
    '====获取区域====
    Dim mapcells
        mapcells = Array(0, 1, 3, 5, 7, 9, 11, 13, 14)
End Function
' 克隆数组
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function CloneAry(ByVal ary As Variant) As Variant
    If Not IsArray(ary) Then Exit Function
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
End Function

' 检查两个数组是否相等
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Boolean
Function IsAryEqual(ByVal ary1 As Variant, ByVal Ary2 As Variant) As Boolean
    IsAryEqual = False
    If Not IsArray(ary1) Or Not IsArray(Ary2) Then Exit Function
    If Not UBound(ary1) = UBound(Ary2) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary1)
        If Not ary1(i) = Ary2(i) Then Exit Function
    Next
    IsAryEqual = True
End Function

'*****IO相关函数*****
' 获取FileSystemObject对象
''' @return:Object(Of FileSystemObject)
Function GetFso() As Object
    Set GetFso = CreateObject("Scripting.FileSystemObject")
End Function

Public Function getRegexp() As Object
    Dim regEx: Set regEx = CreateObject("VBScript.RegExp")
    Set getRegexp = regEx
End Function
Public Function getshell()
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    Set getshell = shellApp
End Function
' 创建Scripting.Dictionary对象
''' @param:CompareMode-Long
''' @return:Object(Of Dictionary)
Function InitDic(Optional compareMode As Long = vbBinaryCompare) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = compareMode
    Set InitDic = dic
End Function
' 创建ArrayList对象
''' @return:Object(Of ArrayList)Public
Function Initlst() As Object
    Set Initlst = CreateObject("System.Collections.ArrayList")
End Function
' 分割路径名
''' @param:FullPath-完整路径
''' @return:Variant(Of Array(Of String)) (0-路径 1-文件名 2-扩展名)
Function SplitPathName(ByVal fullPath$) As Variant
    Dim path(2) As String
    With GetFso
        path(0) = .GetParentFolderName(fullPath)
        path(1) = .GetBaseName(fullPath)
        path(2) = .GetExtensionName(fullPath)
    End With
    SplitPathName = path
End Function
' 合并路径名
''' @param:Path-Variant(Of Array(Of String)) (0-路径 1-文件名 2-扩展名)
''' @return:完整路径
Function JoinPathName$(ByVal path As Variant)
    If Not IsArray(path) Then Stop ' 输入错误
    If Not UBound(path) = 2 Then Stop ' 输入错误
    JoinPathName = path(0) + "\" + path(1) + "." + path(2)
End Function
' 检查路径是否存在
''' @param:Path-路径
''' @return:Boolean
Function isExists(ByVal path$) As Boolean
    isExists = False
    Dim fso As Object: Set fso = GetFso
    If fso.FileExists(path) Then
        isExists = True: Exit Function ' 文件
    ElseIf fso.FolderExists(path) Then
        isExists = True: Exit Function ' 文件夹
    End If
    Set fso = Nothing
End Function
Function GetPath(ByVal path$) As String
    GetPath = ""
     Dim fso As Object: Set fso = GetFso
     If isExists(path) Then
         GetPath = path
     Else
         GetPath = fso.CreateFolder(path)
     End If
     Set fso = Nothing
End Function
'打开并定位到
Sub explorepath(ByVal ipath)
 Dim thisdir, shell, cmd
    thisdir = ""
    Dim fso As Object: Set fso = GetFso
        If fso.FileExists(ipath) Then
            thisdir = ipath
        ElseIf fso.FolderExists(ipath) Then
          Dim Fdl, file
            Set Fdl = fso.GetFolder(ipath)
            For Each file In Fdl.Files
                thisdir = file.path
            Exit For
            Next
        End If
    If thisdir <> "" Then
         Set shell = CreateObject("WScript.Shell")
         cmd = "explorer.exe /select, """ & thisdir & """"
         shell.Run (cmd)
    End If
    Set fso = Nothing
    Set shell = Nothing
End Sub
'获取用户选择路径
Public Function selFdl()
    selFdl = ""
    Dim shellApp, Fdl
    Set shellApp = CreateObject("Shell.Application")
    Set Fdl = shellApp.BrowseForFolder(0, "选择文件夹", 16, 0)
    
    If Not Fdl Is Nothing Then
        selFdl = Fdl.Self.path
    End If
End Function
'@@param: oPath-路径
'获取输入路径父级
Public Function ofParentPath(ByVal opath$)
    Dim idx
    idx = InStrRev(opath, "\")
If idx > 0 Then
        ofParentPath = Left(opath, idx)
    Else
        ofParentPath = opath
    End If
End Function

Sub ClearDir(folderPath As String)
    Dim fso As Object::  Set fso = GetFso()
    ' 检查目录是否存在
    If fso.FolderExists(folderPath) Then
        Dim folder As Object
        Set folder = fso.GetFolder(folderPath)
        ' 删除目录中的所有文件
        Dim file As Object
        For Each file In folder.Files
            fso.DeleteFile file.path, True ' True表示强制删除只读文件
        Next
    End If
    Set fso = Nothing     ' 释放对象
End Sub
Function DeleteMe(ByVal path$) As Boolean
DeleteMe = False
On Error Resume Next
    Dim fso As Object: Set fso = GetFso
    If fso.FileExists(path) Then
        fso.DeleteFile path, True
        DeleteMe = True
    End If
    If Error.Number = 0 Then
        DeleteMe = True
    Else
        Error.Clear
    End If
    Set fso = Nothing
    Error.Clear
On Error GoTo 0
End Function
' 获取新文件名
''' @param:Path-完整路径
''' @return:新的完整路径
Function GetNewName$(ByVal oldPath$)
    Dim path As Variant
    path = SplitPathName(oldPath)
    path(2) = "." & path(2)
    Dim newPath$: newPath = path(0) + "\" + path(1)
    If Not isExists(newPath + path(2)) Then
        GetNewName = newPath + path(2)
        Exit Function
    End If
    Dim tempName$, i&: i = 0
    Do
        i = i + 1
        tempName = newPath + "_" + CStr(i) + path(2)
        If Not isExists(tempName) Then
            GetNewName = tempName
            Exit Function
        End If
    Loop
End Function
' 写入文件
''' @param:Path-完整路径
''' @param:Txt-String
Sub WriteFile(ByVal path$, ByVal txt) '$)
    On Error Resume Next
        Call GetFso.OpenTextFile(path, 2, True).Write(txt)
    On Error GoTo 0
End Sub
' 读取文件
''' @param:Path-完整路径
''' @return:Variant(Of Array(Of String))
Function ReadFile(ByVal path$) As Variant
    On Error Resume Next
    With GetFso.GetFile(path).OpenAsTextStream
        ReadFile = Split(.ReadAll, vbNewLine)
        .Close
    End With
    On Error GoTo 0
End Function
Public Function GetInput(msg) As String
    Dim UserInput As String
    UserInput = InputBox(msg, "输入提示")
    If UserInput = "" Or UserInput = "0" Then
        GetInput = ""
    Else
        GetInput = UserInput
    End If
End Function
' 检查字符串中是否包含指定关键字
' 忽略大小写进行检查
Public Function ExistsKey(ByVal txt As String, ByVal KEY As String) As Boolean
    ExistsKey = IIf(VBA.InStr(VBA.LCase(txt), VBA.LCase(KEY)) > 0, True, False)
End Function
'@@ param:ostr-时间格式
Public Function timestamp(Optional ByVal ostr) As String
    Dim FT As String  ' 显式声明变量
    Select Case True
        Case ExistsKey(ostr, "i"): FT = "yymmdd.hhnn"
        Case ExistsKey(ostr, "h"): FT = "yymmdd.hh"
        Case ExistsKey(ostr, "d"): FT = "yymmdd"
        Case ExistsKey(ostr, "s"): FT = "yymmdd.hhnnss"
        Case Else: FT = "yymmdd"  ' 默认格式，避免未赋值情况
    End Select
    timestamp = Format(Now, FT)
End Function
Function isEngPath(ByVal path As String) As Boolean
    Dim i As Long, charCode As Long
    Dim validChars As String
     validChars = "!@#$%^&*()-_=+[]{};:'"",.<>/?\|~\/"    ' 定义允许的英文符号（包括路径分隔符）
    For i = 1 To Len(path)      ' 遍历路径中的每个字符
        charCode = AscW(Mid(path, i, 1))  ' 检查是否为英文字母（A-Z, a-z）
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            GoTo NextChar  ' 等同于 Continue For
        End If
        If charCode >= 48 And charCode <= 57 Then    ' 检查是否为数字（0-9）
            GoTo NextChar  ' 等同于 Continue For
        End If
        If InStr(validChars, Mid(path, i, 1)) > 0 Then    ' 检查是否为允许的英文符号
            GoTo NextChar  ' 等同于 Continue For
        End If
        isEngPath = False          ' 如果都不是，则路径包含非法字符
        Exit Function
NextChar:
    Next i
    ' 所有字符都通过检查
    isEngPath = True
End Function
' 检查字符串是否包含特殊符号
Function HasSpecialChars(ByVal S As String) As Boolean
    Dim validChars As String
    Dim i As Long
    validChars = "!@#$%^&*()-_=+[]{};:'"",.<>/?\|~\/"
    For i = 1 To Len(S)
        If InStr(validChars, Mid(S, i, 1)) > 0 Then
            HasSpecialChars = True
            Exit Function
        End If
    Next i
    HasSpecialChars = False
End Function

'此函数替换字符串中的特殊符号

Function ReplaceSpcChar(ByVal S As String) As String
    Dim regEx: Set regEx = getRegexp()
    ' 1. Windows系统非法: \ / : * ? " < > |
    ' 2. 用户自定义非法: ! @ # 【 】 { } ~ 《 》 ， ^ % & （ ） ( ) 、
    regEx.Pattern = "[\\/:*?""<>|!@#【】{}~《》，\^%&（）\(\)、]+"
    regEx.Global = True
    ReplaceSpcChar = regEx.Replace(S, "_")
    Set regEx = Nothing
End Function

''替换字符串的所有中文为横线
Function rmchn(ByVal inputString$) As String
    Dim regEx: Set regEx = getRegexp()
    regEx.Pattern = "[\u4e00-\u9fa5]+"
    regEx.Global = True
    rmchn = regEx.Replace(inputString, "_")
    Set regEx = Nothing
End Function

Function ReplaceBadChar(ByVal S As String) As String
    Dim regEx: Set regEx = getRegexp()
    ' 组合模式：
    ' 1. [\\/:*?""<>|!@#【】{}~《》，\^%&（）\(\)、]  -> 特殊符号
    ' 2. \u4e00-\u9fa5                          -> 中文
    ' 3. +                                      -> 连续匹配
    regEx.Pattern = "[\\/:*?""<>|!@#【】Φ{}~《》，\^%&（）\(\)、\u4e00-\u9fa5]+"
    regEx.Global = True
    ReplaceBadChar = regEx.Replace(S, "_")
    Set regEx = Nothing
End Function

' 此函数用于检查输入的路径是否包含中文字符
' 参数: pathToCheck - 需要检查的路径
' 返回值: Boolean 类型，True 表示路径包含中文，False 表示不包含
Function isPathchn(pathToCheck) As Boolean
    Dim regEx As Object
    Set regEx = getRegexp
    regEx.Pattern = "[\u4e00-\u9fa5]"   ' 设置正则表达式模式，匹配中文字符
    regEx.IgnoreCase = True
    regEx.Global = True
    isPathchn = regEx.Test(pathToCheck)   ' 执行匹配并返回结果
    Set regEx = Nothing
End Function


'@iStr string
'获得字符串最后一个"ikey"之前的字符或返回原字符
Function strbflast(str, ikey)
Dim idx
idx = InStrRev(str, ikey)
If idx > 0 Then
        strbflast = Left(str, idx)
    Else
        strbflast = str
    End If
End Function

'@iStr string 原字符串
'@ikey string 分隔符
'返回：第一个ikey之前的字符，未找到则返回原字符串
Function StrBF(istr, ikey)
    Dim pos As Long
    pos = InStr(istr, ikey)
    StrBF = IIf(pos > 0, Left(istr, pos - 1), istr)
End Function
 '@iStr string
'获得字符串第一个"ikey"之后的内容，若无则原内容
Function StrAF(istr, ikey)
   Dim pos As Long
    pos = InStr(istr, ikey)
    StrAF = IIf(pos > 0, Mid(istr, pos + Len(ikey)), istr)
End Function
'创建md文件
Function getmd(ByVal ipath_name As String)
     Dim fso
    Set fso = GetFso()
    Dim mdfile
    If Not KCL.isExists(ipath_name) Then
        Set mdfile = fso.CreateTextFile(ipath_name, False) '不存在则创建
    Else
        Set mdfile = fso.OpenTextFile(ipath_name, ForAppending, TristateFalse) '存在则
    End If
    Set getmd = mdfile
    Set mdfile = Nothing

End Function
'文本文件写入
Sub Appendtext(ByVal tfile As Object, _
           ByVal iText$ _
           )
    tfile.WriteLine (iText)
    Set tfile = Nothing
End Sub
' 获取CATIA 程序语言
Function GetLanguage() As String
    GetLanguage = "non"
    If CATIA.Windows.count < 1 Then Exit Function
    GetLanguage = "other"
    CATIA.ActiveDocument.Selection.Clear
    Dim st As String: st = CATIA.StatusBar
    Select Case True
        Case ExistsKey(st, "object"): GetLanguage = "en"
        Case ExistsKey(st, "objet"):: GetLanguage = "fr"
        Case ExistsKey(st, "Objekt"): GetLanguage = "de"
        Case ExistsKey(st, "oggetto"): GetLanguage = "it"
        Case ExistsKey(st, "命令"):    GetLanguage = "ja"
        Case ExistsKey(st, "объект"):  GetLanguage = "ru"
        Case ExistsKey(st, "对象"): GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13: GetLanguage = "ko"
                Case 23: GetLanguage = "ja"
                Case Else            ' 其他情况
            End Select
    End Select
End Function

Function getVbaDir() As String
    Dim oApc As Object
    Set oApc = GetApc()
    Dim projFilePath As String
    projFilePath = oApc.ExecutingProject.VBProject.fileName
     getVbaDir = GetFso.GetParentFolderName(projFilePath)
End Function
'******* APC/VBE *********
' 获取APC对象
' 参数  :
' 返回值: obj-IApc
Public Function GetApc() As Object
    Set GetApc = Nothing
    Dim COMObjectName$     ' 获取VBA版本对应的COM对象名称
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        MsgBox "不支持当前VBA版本", vbExclamation + vbOKOnly
        Exit Function
    #End If
    Dim Apc As Object: Set Apc = Nothing   ' 获取APC对象
    On Error Resume Next
        Set Apc = CreateObject(COMObjectName)
    On Error GoTo 0
    If Apc Is Nothing Then
        MsgBox "无法获取MSAPC.Apc对象", vbExclamation + vbOKOnly
        Exit Function
    End If
    Set GetApc = Apc
End Function
Function GetVBAprj() As String
    Dim oApc As Object
    Set oApc = GetApc()
    GetVBAprj = oApc.ExecutingProject.VBProject.fileName
End Function
' 智能打开路径（优先激活已存在窗口）
Sub SmartOPenPath(ByVal strPath As String)
    Dim fso As Object: Set fso = GetFso
    If Len(strPath) > 3 And Right(strPath, 1) = "\" Then
        strPath = Left(strPath, Len(strPath) - 1)
    End If
    If Not (fso.FileExists(strPath) Or fso.FolderExists(strPath)) Then
        MsgBox "路径不存在: " & strPath, vbExclamation
        Exit Sub
    End If
    If Not ActivateExistingWindow(strPath) Then     ' 尝试激活已存在的窗口
        If fso.FileExists(strPath) Then  ' 未找到已存在窗口，执行新打开操作
            OpenFileLocation strPath
        ElseIf fso.FolderExists(strPath) Then
            openDir strPath
        End If
    End If
End Sub

' 检查并激活已存在的窗口
Function ActivateExistingWindow(ByVal strPath As String) As Boolean
    Dim W As Object
    On Error Resume Next
    For Each W In CreateObject("Shell.Application").Windows
        If LCase(W.Document.folder.Self.path) = LCase(strPath) Then
            If Err.Number = 0 Then ' 确保路径访问没报错
                 ShowWindow W.hwnd, 1        ' 1 = SW_SHOWNORMAL (普通模式/还原)
                SetForegroundWindow W.hwnd  ' 激活到前台
                ActivateExistingWindow = True
                Exit Function
            End If
            Err.Clear
        End If
    Next
End Function

' 打开文件夹
Private Sub openDir(ByVal strPath As String)
    On Error GoTo ErrorHandler
    If InStr(strPath, " ") > 0 Then      ' 处理包含空格的路径
        strPath = """" & strPath & """"
    End If
    shell "explorer.exe " & strPath, vbMaximizedFocus
    Exit Sub
ErrorHandler:
    MsgBox "无法打开路径: " & strPath & vbCrLf & "错误: " & Err.Description, vbExclamation
End Sub

' 打开文件位置并选中文件
Private Sub OpenFileLocation(ByVal strFilePath As String)
    On Error GoTo ErrorHandler
    strFilePath = """" & strFilePath & """"  ' 确保文件路径被引号包围
    shell "explorer.exe /select," & strFilePath, vbMaximizedFocus
    Exit Sub
ErrorHandler:
    MsgBox "无法打开文件位置: " & strFilePath & vbCrLf & "错误: " & Err.Description, vbExclamation
End Sub

' 批量打开多个路径
'Sub OpenMultiple(ParamArray Paths() As Variant)
'    Dim i As Long
'    For i = LBound(Paths) To UBound(Paths)
'        SmartOpen CStr(Paths(i))
'        DoEvents ' 允许系统处理其他事件
'    Next i
'End Sub
Public Function Push_Dic(ByVal dic As Object, _
                          ByVal KEY As Variant, _
                          ByVal item As Variant) As Object
    If dic.Exists(KEY) Then
        dic(KEY) = item
    Else
        dic.Add KEY, item
    End If
    Set Push_Dic = dic
End Function

Public Function showdict(ByVal oDic, Optional ByVal boolShowKeyIndex As Boolean = False)
  Dim keys:   keys = oDic.keys
  Dim i As Long
  Dim stIndex As String
  Dim stOutput As String
  stOutput = vbNullString
  
  For i = 0 To oDic.count - 1
    If boolShowKeyIndex Then
      stIndex = "(" & i & ")"
    End If
    stOutput = stOutput & keys(i) & stIndex & "  :  "
    If IsObject(oDic(keys(i))) Then
      stOutput = stOutput & "[" & showdict(oDic(keys(i)), boolShowKeyIndex) & "]"
    Else
      stOutput = stOutput & oDic(keys(i))
    End If
        stOutput = stOutput & "; " & "_" & vbNewLine
  Next i
  showdict = stOutput
  Debug.Print showdict
End Function

' 获取Brep名称
''' @param:MyBRepName-String
''' @return:String
Public Function GetBrepName(MyBRepName As String) As String
    MyBRepName = Replace(MyBRepName, "Selection_", "")
    MyBRepName = Left(MyBRepName, InStrRev(MyBRepName, "));"))
    MyBRepName = MyBRepName + ");WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    GetBrepName = MyBRepName
End Function
Function getmdl(Optional ByVal modName As String = "")
  Set getmdl = Nothing
    Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    On Error Resume Next
     Dim mdl:
     If modName = "" Then
        Set mdl = ExecPjt.VBProject.VBE.Activecodepane.CodeModule
     Else
        Set mdl = ExecPjt.VBProject.VBComponents(modName)
    End If
        Error.Clear
    On Error GoTo 0
    If mdl Is Nothing Then Exit Function
   Set getmdl = mdl
End Function

Public Function getDecCode(modName)
 Dim DecCnt
   Dim mdl:  Set mdl = KCL.getmdl(modName)
    If mdl Is Nothing Then Exit Function
        DecCnt = mdl.CodeModule.CountOfDeclarationLines ' 获取声明行数
    If DecCnt < 1 Then Exit Function
        getDecCode = mdl.CodeModule.lines(1, DecCnt) ' 获取声明代码
End Function

Public Function getbf1stproc(modName)
    getbf1stproc = ""
    Dim mdl:  Set mdl = KCL.getmdl(modName)
    Dim codemod, i As Long, procName, startline As Long
    Dim prockind As Long
    If mdl Is Nothing Then Exit Function
    Set codemod = mdl.CodeModule
    For i = 1 To codemod.CountOfLines
        procName = codemod.ProcOfLine(i, prockind)
        If procName <> "" Then startline = codemod.ProcBodyLine(procName, prockind)  ' procKind是枚举:sub/function/property
    Next
    If startline < 1 Then Exit Function
        getbf1stproc = codemod.lines(1, startline) ' 获取到第一个函数的所有代码行
End Function

Function getmeas(itm)
    Set getmeas = Nothing
   If Not itm Is Nothing Then
       Dim oDoc: Set oDoc = CATIA.ActiveDocument
      Dim spa:  Set spa = oDoc.GetWorkbench("SPAWorkbench")
        Set getmeas = spa.GetMeasurable(itm)
    End If
End Function
'Function setBTNmdl(ByVal modName As String)
'    Set setBTNmdl = Nothing
'    Dim ctrllst:    Set ctrllst = KCL.ParseUIConfig(KCL.getbf1stproc(modName))
'    Dim map: Set map = KCL.InitDic
'    Dim oCtrl
'    For Each oCtrl In ctrllst    '映射BTN名字和对应模块
'        Select Case oCtrl("Type")
'            Case "Forms.CommandButton.1"
'                map(oCtrl("Name")) = modName
'        End Select
'    Next
'   Set setBTNmdl = map
'End Function
'Function setBTNFunc(ByVal modName As String)
'    Set setBTNFunc = Nothing
'    Dim ctrllst:    Set ctrllst = KCL.ParseUIConfig(KCL.getbf1stproc(modName))
'    Dim map: Set map = KCL.InitDic
'    Dim oCtrl
'    For Each oCtrl In ctrllst    '映射BTN名字和对应函数
'        Select Case oCtrl("Type")
'            Case "Forms.CommandButton.1"
'                map(oCtrl("Name")) = oCtrl("Name") & "_Click"
'        End Select
'    Next
'   Set setBTNFunc = mapssss
'End Function
'Function newFrm(Optional ByVal modName As String = "", Optional ByVal isVertical = False)
'    Dim oFrm: Set oFrm = New cls_dynaFrm
'    If modName <> "" Then
'        oFrm.Init modName
'   End If
'   If isVertical Then oFrm.isVertical = True
'   Set newFrm = oFrm
'End Function
' ═══ 新的统一UI引擎工厂函数 ═══
' 用法与 newFrm 平行，返回 Cls_DynaUIEngine 实例
' 模态弹窗示例:  Set oEng = KCL.newEngine("OTH_Minibox") : oEng.Show
' 工具栏示例:    Set oEng = KCL.newEngine("OTH_ivhideshow") : oEng.ShowToolbar mdlname, mapMdl, mapFunc
' 纯代码示例:    Set oEng = KCL.newEngine() : oEng.AddUIElement "Button","btn1","确定" : oEng.Show
Function newEngine(Optional ByVal modName As String = "", Optional ByVal isVertical = False)
    Dim oEng As Cls_DynaUIEngine: Set oEng = New Cls_DynaUIEngine
    If modName <> "" Then
        oEng.LoadFromModuleName modName
    End If
    If isVertical Then oEng.isVertical = True
    Set newEngine = oEng
End Function
Public Function ParseHex(ByVal hexStr)
Dim R, G, B
    hexStr = Replace(hexStr, "#", "")
    If Len(hexStr) <> 6 Then
        ParseHex = ""
        Exit Function
    End If
    On Error Resume Next ' 防止非法字符报错
    R = val("&H" & Mid(hexStr, 1, 2))
    G = val("&H" & Mid(hexStr, 3, 2))
    B = val("&H" & Mid(hexStr, 5, 2))
   
    On Error GoTo 0
   ParseHex = RGB(R, G, B)
End Function

Public Function ParseBDcolor(ByVal hexStr)
Dim R, G, B
    hexStr = Replace(hexStr, "#", "")
    If Len(hexStr) <> 6 Then
        ParseBDcolor = ""
        Exit Function
    End If
    On Error Resume Next ' 防止非法字符报错
    R = val("&H" & Mid(hexStr, 1, 2))
    G = val("&H" & Mid(hexStr, 3, 2))
    B = val("&H" & Mid(hexStr, 5, 2))
    On Error GoTo 0
   ParseBDcolor = Array(R, G, B)
End Function

' --- UI Parsing Helpers ---'
'返回一个列表对象，每个元素是字典，包含：
'Type：控件类型（经过 GetControlType 转换）
'Name: 控件名称
'Caption: 显示文字
'Color: 颜色 (可选)
' --- UI Parsing Helpers ---'
Function ParseUIConfig(ByVal code As String) As Object
    Dim regEx As Object, matches As Object, match As Object
    Dim CtrCfg As Object, cfglst As Object
    Set cfglst = KCL.Initlst
    Set regEx = KCL.getRegexp
    With regEx
        .Global = True
        .MultiLine = True
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.+?)(?:\s+(#[0-9a-fA-F]{6}))?\s*$"
        '      "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.*)$"
    End With
    If regEx.Test(code) Then
        Set matches = regEx.Execute(code)
        For Each match In matches
            Set CtrCfg = KCL.InitDic
            CtrCfg.Add "Type", GetControlType(match.SubMatches(0))
            CtrCfg.Add "Name", match.SubMatches(1)
            CtrCfg.Add "Caption", VBA.Trim(match.SubMatches(2))
            CtrCfg.Add "Color", VBA.Trim(match.SubMatches(3))
            Debug.Print CtrCfg("Color")
            cfglst.Add CtrCfg
        Next
    End If
    Set ParseUIConfig = cfglst
End Function

Function ParseUITitle(ByVal code As String) As String
    Dim regEx As Object, matches As Object, match As Object, itl
    Set regEx = KCL.getRegexp
    With regEx
        .Global = True
        .MultiLine = True
        .Pattern = "^\s*'\s*%Title\s+(.*)$"
    End With
    itl = "请问如何执行"
    If regEx.Test(code) Then
        Set matches = regEx.Execute(code)
        For Each match In matches
            itl = match.SubMatches(0)
        Next
    End If
    ParseUITitle = itl
End Function

Function GetControlType(ByVal alias As String) As String
    alias = LCase(alias)
    Select Case alias
        Case "button", "btn", "cmd": GetControlType = "Forms.CommandButton.1"
        Case "label", "lbl": GetControlType = "Forms.Label.1"
        Case "text", "txt", "textbox": GetControlType = "Forms.TextBox.1"
        Case "checkbox", "chk": GetControlType = "Forms.CheckBox.1"
        Case Else: GetControlType = "Forms.Label.1"
    End Select
End Function
Public Function CATquick(ByVal Quick As Boolean, Optional ByVal updateCap As Boolean = False)
    Dim setcls:  Set setcls = CATIA.SettingControllers
    Dim Asmg:   Set Asmg = setcls.item("CATAsmGeneralSettingCtrl")
    Dim Vismg:   Set Vismg = setcls.item("CATVizVisualizationSettingCtrl")
    Dim btnCaption As String
    With CATIA
    
    If Quick Then
        '.DisableNewUndoRedoTransaction
        '.EnableNewUndoRedoTransaction
         .RefreshDisplay = False
            Asmg.AutoUpdateMode = 0 '0: catManualUpdate
            Vismg.Viz3DFixedAccuracy = 5
            btnCaption = "屏幕更新(关)"
    Else
        '.DisableNewUndoRedoTransaction
        '.EnableNewUndoRedoTransaction
        .RefreshDisplay = True
       Asmg.AutoUpdateMode = 1 '1: catAutomaticUpdate
        Vismg.Viz3DFixedAccuracy = 0.02
        btnCaption = "屏幕更新(开)"
    End If
    End With
     On Error Resume Next
        If Not g_Btn Is Nothing Then
          If updateCap = True Then g_Btn.Caption = btnCaption
        End If
     Set g_Btn = Nothing
    CATquick = Quick
    CATIA.ActiveWindow.ActiveViewer.Update
    On Error GoTo 0
End Function





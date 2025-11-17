Attribute VB_Name = "KCL"
'Attribute VB_Name = "KCL"
'vba Kantoku_CATVBA_Library ver0.1.0
'KCL.bas - 自定义VBA库
Option Explicit

Private mSW& ' 秒表开始时间

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

' 主程序入口 - 循环选择项目
Sub CATMain()
    Dim msg$: msg = "请选择项目 : 按ESC键退出"
    Dim SI As AnyObject
    Dim doc As Document: Set doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(msg)
        If IsNothing(SI) Then Exit Do
        Stop
    Loop
End Sub
'*****CATIA相关函数*****
' 检查是否可以执行操作
''' @param:DocTypes-array(string),string 指定可执行操作的文档类型
''' @return:Boolean
Function CanExecute(ByVal docTypes As Variant) As Boolean
    CanExecute = False
    If CATIA.Windows.count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Function
    End If
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",") '过滤器转数组
    If Not checkFilterType(docTypes) Then Exit Function '过滤器检查，非数组则退出
    Dim ErrMsg As String
    ErrMsg = "不支持当前活动文档类型。" + vbNewLine + "(" + Join(docTypes, ",") + " 类型除外)"
    CanExecute = checkDocType(docTypes)
    If Not CanExecute Then MsgBox ErrMsg, vbExclamation + vbOKOnly
End Function
Function checkDocType(ByVal docTypes As Variant)
    checkDocType = False
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",") '过滤器转数组
    If Not checkFilterType(docTypes) Then Exit Function '过滤器检查，非数组则退出
    Dim ActDoc As Document
    On Error Resume Next
        Set ActDoc = CATIA.ActiveDocument
    On Error GoTo 0
    If ActDoc Is Nothing Then
        MsgBox "无打开的文档"
        Exit Function
    End If
     If UBound(filter(docTypes, TypeName(ActDoc))) < 0 Then '此处filter函数是VBA中比较厚返回数组的函数
        Exit Function
    End If
    checkDocType = True
End Function
' 选择项目
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
        Set SelectItem = se.value
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
    Dim sel As Variant: Set sel = CATIA.ActiveDocument.Selection
    sel.Clear
    Select Case sel.SelectElement2(filter, msg, False)
        Case "Cancel", "Undo", "Redo"
             Set SelectElement = Nothing
            Exit Function
    End Select
    Set SelectElement = sel.item(1)
    sel.Clear
End Function
' 获取内部名称
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(aoj)
    If IsNothing(aoj) Then
        GetInternalName = Empty: Exit Function
    End If
    GetInternalName = aoj.GetItem("ModelElement").InternalName
End Function

' 获取指定类型的父对象
''' @param:anyObj-AnyObject
''' @param:T-String
''' @return:AnyObject
Function GetParent_Of_T( _
                        ByVal anyObj As AnyObject, _
                        ByVal t As String) As AnyObject
    Dim anyObjName As String
    Dim parentName As String
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
' 获取Brep名称
''' @param:MyBRepName-String
''' @return:String
Function GetBrepName(MyBRepName As String) As String
    MyBRepName = Replace(MyBRepName, "Selection_", "")
    MyBRepName = Left(MyBRepName, InStrRev(MyBRepName, "));"))
    MyBRepName = MyBRepName + ");WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    GetBrepName = MyBRepName
End Function
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
End Function ' 检查是否为字符串数组
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    IsStringAry = True
End Function
' 将字符串转换为变体数组
Private Function strToAry(ByVal s$) As Variant
    Dim ary As Variant: ary = Split(s, ",")
    
    Dim oAry() As Variant: ReDim oAry(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        oAry(i) = ary(i)
    Next
    strToAry = oAry
End Function
' 检查过滤器类型是否有效
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
'*****通用相关函数*****
' 检查对象是否为Nothing
''' @param:OJ-Variant(Of Object)
''' @return:Boolean
Function IsNothing(ByVal oj As Variant) As Boolean
    IsNothing = oj Is Nothing
End Function
' 创建Scripting.Dictionary对象
''' @param:CompareMode-Long
''' @return:Object(Of Dictionary)
Function InitDic(Optional compareMode As Long = vbBinaryCompare) As Object
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.compareMode = compareMode
    Set InitDic = Dic
End Function
' 创建ArrayList对象
''' @return:Object(Of ArrayList)Public
Function InitLst() As Object
    Set InitLst = CreateObject("System.Collections.ArrayList")
End Function
' 检查对象是否为指定类型
''' @param:OJ-Object
''' @param:T-String
''' @return:Boolean
Function isobjtype(ByVal oj As Object, ByVal t$) As Boolean
    isobjtype = IIf(TypeName(oj) = t, True, False)
'    MsgBox TypeName(oj)
End Function

'*****数组相关函数*****
' 合并两个数组
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Variant(Of Array)
Function JoinAry(ByVal ary1 As Variant, ByVal ary2 As Variant)
    Select Case True
        Case Not IsArray(ary1) And Not IsArray(ary2)
            JoinAry = Empty: Exit Function
        Case Not IsArray(ary1)
            JoinAry = ary2: Exit Function
        Case Not IsArray(ary2)
            JoinAry = ary1: Exit Function
    End Select
    Dim StCount&: StCount = UBound(ary1)
    ReDim Preserve ary1(UBound(ary1) + UBound(ary2) + 1)
    Dim i&
    If IsObject(ary2(0)) Then
        For i = StCount + 1 To UBound(ary1)
            Set ary1(i) = ary2(i - StCount - 1)
        Next
    Else
        For i = StCount + 1 To UBound(ary1)
            ary1(i) = ary2(i - StCount - 1)
        Next
    End If
    JoinAry = ary1
End Function
' mapping 数组
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function mappedAry(ByVal ary As Variant, ByVal iMap As Variant) As Variant
    If Not IsArray(ary) Or Not IsArray(iMap) Then Exit Function
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
    Dim ele_Ary()
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
Function IsAryEqual(ByVal ary1 As Variant, ByVal ary2 As Variant) As Boolean
    IsAryEqual = False
    If Not IsArray(ary1) Or Not IsArray(ary2) Then Exit Function
    If Not UBound(ary1) = UBound(ary2) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary1)
        If Not ary1(i) = ary2(i) Then Exit Function
    Next
    IsAryEqual = True
End Function

'*****IO相关函数*****
' 获取FileSystemObject对象
''' @return:Object(Of FileSystemObject)
Function GetFso() As Object
    Set GetFso = CreateObject("Scripting.FileSystemObject")
End Function

' 分割路径名
''' @param:FullPath-完整路径
''' @return:Variant(Of Array(Of String)) (0-路径 1-文件名 2-扩展名)
Function SplitPathName(ByVal fullpath$) As Variant
    Dim path(2) As String
    With GetFso
        path(0) = .GetParentFolderName(fullpath)
        path(1) = .GetBaseName(fullpath)
        path(2) = .GetExtensionName(fullpath)
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
    Dim FSO As Object: Set FSO = GetFso
    If FSO.FileExists(path) Then
        isExists = True: Exit Function ' 文件
    ElseIf FSO.FolderExists(path) Then
        isExists = True: Exit Function ' 文件夹
    End If
    Set FSO = Nothing
End Function
Function GetPath(ByVal path$)
    GetPath = ""
     Dim FSO As Object: Set FSO = GetFso
     If isExists(path) Then
         GetPath = path
     Else
         GetPath = FSO.CreateFolder(path)
     End If
     Set FSO = Nothing
End Function

Sub explorepath(ByVal ipath)
 Dim thisdir, shell, cmd
    thisdir = ""
    Dim FSO As Object: Set FSO = GetFso
        If FSO.FileExists(ipath) Then
            thisdir = ipath
        ElseIf FSO.FolderExists(ipath) Then
        Dim Fdl, file
            Set Fdl = FSO.GetFolder(ipath)
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
    Set FSO = Nothing
    Set shell = Nothing
End Sub

Public Function selFdl()
    selFdl = ""
    Dim shellApp, Fdl
    Set shellApp = CreateObject("Shell.Application")
    Set Fdl = shellApp.BrowseForFolder(0, "选择文件夹", 16, 0)
    
    If Not Fdl Is Nothing Then
        selFdl = Fdl.Self.path
    End If
End Function



Sub ClearDir(folderPath As String)
    Dim FSO As Object
    Set FSO = GetFso()
    ' 检查目录是否存在
    If FSO.FolderExists(folderPath) Then
        Dim folder As Object
        Set folder = FSO.GetFolder(folderPath)
        ' 删除目录中的所有文件
        Dim file As Object
        For Each file In folder.Files
            FSO.DeleteFile file.path, True ' True表示强制删除只读文件
        Next
    End If
    Set FSO = Nothing     ' 释放对象
End Sub

Function DeleteMe(ByVal path$) As Boolean
DeleteMe = False
On Error Resume Next
    Dim FSO As Object: Set FSO = GetFso
    
    If FSO.FileExists(path) Then
        FSO.DeleteFile path, True
        DeleteMe = True
    End If
    
    If Error.Number = 0 Then
        DeleteMe = True
    Else
        Error.Clear
    End If
    
    Set FSO = Nothing
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
Sub WriteFile(ByVal path$, ByVal Txt) '$)
    On Error Resume Next
        Call GetFso.OpenTextFile(path, 2, True).Write(Txt)
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
'*****计时相关函数*****
' 启动秒表
Sub SW_Start()
    mSW = timeGetTime
End Sub

' 获取计时时间
''' @return:Double(Unit:s)
Function SW_GetTime#()
    SW_GetTime = IIf(mSW = 0, -1, (timeGetTime - mSW) * 0.001)
End Function

Public Function GetInput(msg) As String
    Dim UserInput As String
    UserInput = InputBox(msg, "输入提示")
    ' 如果用户没有输入或点击取消，则返回默认值"XX"
    If UserInput = "" Or UserInput = "0" Then
        GetInput = ""
    Else
        GetInput = UserInput
    End If
End Function

'@@param: oPath-路径
'获取输入路径父级
Public Function ofParentPath(ByVal oPath$)
    Dim idx
    idx = InStrRev(oPath, "\")
If idx > 0 Then
        ofParentPath = Left(oPath, idx)
    Else
        ofParentPath = oPath
    End If
End Function
' 检查字符串中是否包含指定关键字
' 忽略大小写进行检查
Public Function ExistsKey(ByVal Txt As String, ByVal Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(Txt), LCase(Key)) > 0, True, False)
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
    ' 定义允许的英文符号（包括路径分隔符）
    validChars = "!@#$%^&*()-_=+[]{};:'"",.<>/?\|~\/"
    ' 遍历路径中的每个字符
    For i = 1 To Len(path)
        charCode = AscW(Mid(path, i, 1))
        ' 检查是否为英文字母（A-Z, a-z）
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            GoTo NextChar  ' 等同于 Continue For
        End If
        ' 检查是否为数字（0-9）
        If charCode >= 48 And charCode <= 57 Then
            GoTo NextChar  ' 等同于 Continue For
        End If
        ' 检查是否为允许的英文符号
        If InStr(validChars, Mid(path, i, 1)) > 0 Then
            GoTo NextChar  ' 等同于 Continue For
        End If
        ' 如果都不是，则路径包含非法字符
        isEngPath = False
        Exit Function
NextChar:
    Next i
    ' 所有字符都通过检查
    isEngPath = True
End Function

' 此函数用于检查输入的路径是否包含中文字符
' 参数:
'   pathToCheck - 需要检查的路径
' 返回值:
'   Boolean 类型，True 表示路径包含中文，False 表示不包含
Function isPathchn(pathToCheck) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 设置正则表达式模式，匹配中文字符
    regex.Pattern = "[\u4e00-\u9fa5]"
    regex.IgnoreCase = True
    regex.Global = True
    ' 执行匹配并返回结果
    isPathchn = regex.test(pathToCheck)
    Set regex = Nothing
End Function
'@iStr string
'获得字符串最后一个"iext"之前的字符或返回原字符
Function strbflast(str, iext)
Dim idx
idx = InStrRev(str, iext)
If idx > 0 Then
        strbflast = Left(str, idx)
    Else
        strbflast = str
    End If
End Function

'@iStr string
'获得字符串第一个"_"之前的字符或返回原字符
Function strbf1st(iStr, iext)
    Dim oPrefix
        Dim underscorePos As Long
        underscorePos = InStr(iStr, iext)
        If underscorePos > 0 Then
            oPrefix = Left(iStr, underscorePos - 1)
        Else
           oPrefix = iStr
        End If
        strbf1st = oPrefix
End Function
 
Function straf1st(iStr, iext)
Dim idx
idx = InStr(iStr, iext)
If idx > 0 Then
        straf1st = Mid(iStr, idx)
    Else
        straf1st = iStr
    End If
End Function

''替换字符串的所有中文为空格
Function rmchn(ByVal inputString$) As String
    Dim regex: Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "[\u4e00-\u9fa5]"
    regex.Global = True
    rmchn = regex.Replace(inputString, " ")
    Set regex = Nothing
End Function


'创建md文件
Function getmd(ByVal ipath_name As String)
     Dim FSO
    Set FSO = GetFso()
    Dim mdfile
    If Not KCL.isExists(ipath_name) Then
        Set mdfile = FSO.CreateTextFile(ipath_name, False) '不存在则创建
    Else
        Set mdfile = FSO.OpenTextFile(ipath_name, ForAppending, TristateFalse) '存在则
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


' 获取语言
'return-ISO 639-1 code
'https://ja.wikipedia.org/wiki/ISO_639-1%E3%82%B3%E3%83%BC%E3%83%89%E4%B8%80%E8%A6%A7
Function GetLanguage() As String
    GetLanguage = "non"
    If CATIA.Windows.count < 1 Then Exit Function
    GetLanguage = "other"
    CATIA.ActiveDocument.Selection.Clear
    Dim st As String: st = CATIA.StatusBar
    Select Case True
        Case ExistsKey(st, "object")
            GetLanguage = "en"
        Case ExistsKey(st, "objet")
            GetLanguage = "fr"
        Case ExistsKey(st, "Objekt")
            GetLanguage = "de"
        Case ExistsKey(st, "oggetto")
            GetLanguage = "it"
        Case ExistsKey(st, "命令")
            GetLanguage = "ja"
        Case ExistsKey(st, "объект")
            GetLanguage = "ru"
        Case ExistsKey(st, "对象")
            GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13
                    GetLanguage = "ko"
                Case 23
                    GetLanguage = "ja"
                Case Else
                    ' 其他情况
            End Select
    End Select
End Function

Function getVbaDir() As String
    Dim oApc As Object
    Set oApc = GetApc()
    Dim projFilePath As String
    projFilePath = oApc.ExecutingProject.VBProject.Filename
     getVbaDir = GetFso.GetParentFolderName(projFilePath)
End Function
Function GetApc() As Object
    Dim COMObjectName As String
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #End If
    Dim oApc As Object
    On Error Resume Next
    Set oApc = CreateObject(COMObjectName)
    On Error GoTo 0
    If oApc Is Nothing Then
        Set oApc = CreateObject("MSAPC.Apc")
    End If
    Set GetApc = oApc
End Function

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
    
    If CATIA.Windows.Count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Function
    End If
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",") '过滤器转数组
    
    If Not checkFilterType(docTypes) Then Exit Function '过滤器检查，非数组则退出
    
    Dim ErrMsg As String
    ErrMsg = "不支持当前活动文档类型。" + vbNewLine + "(" + Join(docTypes, ",") + " 类型除外)"
 
'    Dim ActDoc As Document
'
'    On Error Resume Next
'        Set ActDoc = CATIA.ActiveDocument
'    On Error GoTo 0
'
'    If ActDoc Is Nothing Then
'        MsgBox ErrMsg, vbExclamation + vbOKOnly
'        Exit Function
'    End If
'
'    If UBound(filter(docTypes, TypeName(ActDoc))) < 0 Then 此处filter函数是VBA中比较厚返回数组的函数
'        MsgBox ErrMsg, vbExclamation + vbOKOnly
'        Exit Function
'    End If
    
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
    ByVal t As String) _
    As AnyObject
    
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

' 获取语言
'return-ISO 639-1 code
'https://ja.wikipedia.org/wiki/ISO_639-1%E3%82%B3%E3%83%BC%E3%83%89%E4%B8%80%E8%A6%A7
Function GetLanguage() As String
    GetLanguage = "non"
    If CATIA.Windows.Count < 1 Then Exit Function
    GetLanguage = "other"
    CATIA.ActiveDocument.Selection.Clear
    Dim st As String: st = CATIA.StatusBar
    Select Case True
        Case ExistsKey(st, "object")
            ' 英文-Select an object or a command
            GetLanguage = "en"
        Case ExistsKey(st, "objet")
            ' 法语-选择一个对象或命令
            GetLanguage = "fr"
        Case ExistsKey(st, "Objekt")
            ' 德语-选择一个对象或命令
            GetLanguage = "de"
        Case ExistsKey(st, "oggetto")
            ' 意大利语-选择一个对象或命令
            GetLanguage = "it"
        Case ExistsKey(st, "命令")
            ' 日语-选择一个命令或对象
            GetLanguage = "ja"
        Case ExistsKey(st, "объект")
            ' 俄语-选择一个对象或命令
            GetLanguage = "ru"
        Case ExistsKey(st, "对象")
            ' 中文-选择一个对象或命令
            GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13
                    ' 韩语-???? ?? ?? ??@unicode编码示例
                    GetLanguage = "ko"
                Case 23
                    ' 日语-日语长提示示例
                    GetLanguage = "ja"
                Case Else
                    ' 其他情况
            End Select
    End Select
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

' 将字符串转换为变体数组
Private Function strToAry(ByVal S$) As Variant
    Dim ary As Variant: ary = Split(S, ",")
    
    Dim oAry() As Variant: ReDim oAry(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        oAry(i) = ary(i)
    Next
    
    strToAry = oAry
    
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
Function InitDic(Optional CompareMode As Long = vbBinaryCompare) As Object
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = CompareMode
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
Function IsType_Of_T(ByVal oj As Object, ByVal t$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = t, True, False)
'    MsgBox TypeName(oj)
End Function


'*****数组相关函数*****
' 合并两个数组
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Variant(Of Array)
Function JoinAry(ByVal Ary1 As Variant, ByVal Ary2 As Variant)
    Select Case True
        Case Not IsArray(Ary1) And Not IsArray(Ary2)
            JoinAry = Empty: Exit Function
        Case Not IsArray(Ary1)
            JoinAry = Ary2: Exit Function
        Case Not IsArray(Ary2)
            JoinAry = Ary1: Exit Function
    End Select
    Dim StCount&: StCount = UBound(Ary1)
    ReDim Preserve Ary1(UBound(Ary1) + UBound(Ary2) + 1)
    Dim i&
    If IsObject(Ary2(0)) Then
        For i = StCount + 1 To UBound(Ary1)
            Set Ary1(i) = Ary2(i - StCount - 1)
        Next
    Else
        For i = StCount + 1 To UBound(Ary1)
            Ary1(i) = Ary2(i - StCount - 1)
        Next
    End If
    JoinAry = Ary1
End Function

' 获取数组指定范围的元素
''' @param:Ary-Variant(Of Array)
''' @param:StartIdx-Long
''' @param:EndIdx-Long
''' @return:Variant(Of Array)
Function GetRangeAry(ByVal ary As Variant, ByVal StartIdx&, ByVal EndIdx&) As Variant
    If Not IsArray(ary) Then Exit Function
    If EndIdx - StartIdx < 0 Then Exit Function
    If StartIdx < 0 Then Exit Function
    If EndIdx > UBound(ary) Then Exit Function
    
    Dim RngAry() As Variant: ReDim RngAry(EndIdx - StartIdx)
    Dim i&
    For i = StartIdx To EndIdx
        RngAry(i - StartIdx) = ary(i)
    Next
    GetRangeAry = RngAry
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
Function IsAryEqual(ByVal Ary1 As Variant, ByVal Ary2 As Variant) As Boolean
    IsAryEqual = False
    If Not IsArray(Ary1) Or Not IsArray(Ary2) Then Exit Function
    If Not UBound(Ary1) = UBound(Ary2) Then Exit Function
    Dim i&
    For i = 0 To UBound(Ary1)
        If Not Ary1(i) = Ary2(i) Then Exit Function
    Next
    IsAryEqual = True
End Function


'*****IO相关函数*****
' 获取FileSystemObject对象
''' @return:Object(Of FileSystemObject)
Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

' 分割路径名
''' @param:FullPath-完整路径
''' @return:Variant(Of Array(Of String)) (0-路径 1-文件名 2-扩展名)
Function SplitPathName(ByVal FullPath$) As Variant
    Dim path(2) As String
    With GetFSO
        path(0) = .GetParentFolderName(FullPath)
        path(1) = .GetBaseName(FullPath)
        path(2) = .GetExtensionName(FullPath)
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
    Dim fso As Object: Set fso = GetFSO
    If fso.FileExists(path) Then
        isExists = True: Exit Function ' 文件
    ElseIf fso.FolderExists(path) Then
        isExists = True: Exit Function ' 文件夹
    End If
    Set fso = Nothing
End Function
Function DeleteMe(ByVal path$) As Boolean
    DeleteMe = False
    On Error Resume Next
    Dim fso As Object: Set fso = GetFSO
    If fso.FileExists(path) Then
        fso.DeleteFile path, True
    End If
    If Error.Number = 0 Then
        DeleteMe = True
    Else
    Error.Clear
    End If
    Set fso = Nothing
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
    Dim TempName$, i&: i = 0
    Do
        i = i + 1
        TempName = newPath + "_" + CStr(i) + path(2)
        If Not isExists(TempName) Then
            GetNewName = TempName
            Exit Function
        End If
    Loop
End Function

' 写入文件
''' @param:Path-完整路径
''' @param:Txt-String
Sub WriteFile(ByVal path$, ByVal txt) '$)
    On Error Resume Next
        Call GetFSO.OpenTextFile(path, 2, True).Write(txt)
    On Error GoTo 0
End Sub

' 读取文件
''' @param:Path-完整路径
''' @return:Variant(Of Array(Of String))
Function ReadFile(ByVal path$) As Variant
    On Error Resume Next
    With GetFSO.GetFile(path).OpenAsTextStream
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
Public Function ofParentPath(ByVal opath$)
    Dim idx
    idx = InStrRev(opath, "\")
If idx > 0 Then
        ofParentPath = Left(opath, idx)
    Else
        ofParentPath = opath
    End If
End Function
' 检查字符串中是否包含指定关键字
' 忽略大小写进行检查
Public Function ExistsKey(ByVal txt As String, ByVal Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(Key)) > 0, True, False)
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
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 设置正则表达式模式，匹配中文字符
    regEx.Pattern = "[\u4e00-\u9fa5]"
    regEx.IgnoreCase = True
    regEx.Global = True
    ' 执行匹配并返回结果
    isPathchn = regEx.test(pathToCheck)
    Set regEx = Nothing
End Function
'@iStr string
'获得字符串最后一个"iext"之前的字符或返回原字符
Function strbflast(Str, iext)
Dim idx
idx = InStrRev(Str, iext)
If idx > 0 Then
        strbflast = Left(Str, idx)
    Else
        strbflast = Str
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

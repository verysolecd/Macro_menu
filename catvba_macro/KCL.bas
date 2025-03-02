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
    Dim Msg$: Msg = "请选择项目 : 按ESC键退出"
    Dim SI As AnyObject
    Dim Doc As Document: Set Doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(Msg)
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
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",")
    If Not IsFilterType(docTypes) Then Exit Function
    
    Dim ErrMsg As String
    ErrMsg = "不支持当前活动文档类型。" + vbNewLine + "(" + Join(docTypes, ",") + " 类型除外)"
    
    Dim ActDoc As Document
    On Error Resume Next
        Set ActDoc = CATIA.ActiveDocument
    On Error GoTo 0
    If ActDoc Is Nothing Then
        MsgBox ErrMsg, vbExclamation + vbOKOnly
        Exit Function
    End If
    
    If UBound(Filter(docTypes, TypeName(ActDoc))) < 0 Then
        MsgBox ErrMsg, vbExclamation + vbOKOnly
        Exit Function
    End If
    
    CanExecute = True
End Function

' 选择项目
''' @param:Msg-提示信息
''' @param:Filter-array(string),string 选择过滤器(默认为AnyObject)
''' @return:AnyObject
Function SelectItem(ByVal Msg$, _
                           Optional ByVal Filter As Variant = Empty) _
                           As AnyObject
    Dim SE As SelectedElement
    Set SE = SelectElement(Msg, Filter)
    
    If IsNothing(SE) Then
        Set SelectItem = SE
    Else
        Set SelectItem = SE.Value
    End If
End Function

' 选择元素
''' @param:Msg-提示信息
''' @param:Filter-array(string),string 选择过滤器(默认为AnyObject)
''' @return:SelectedElement
Function SelectElement(ByVal Msg$, _
                           Optional ByVal Filter As Variant = Empty) _
                           As SelectedElement
    If IsEmpty(Filter) Then Filter = Array("AnyObject")
    If VarType(Filter) = vbString Then Filter = ToStrVriAry(Filter)
    If Not IsFilterType(Filter) Then Exit Function
    
    Dim Sel As Variant: Set Sel = CATIA.ActiveDocument.Selection
    Sel.Clear
    Select Case Sel.SelectElement2(Filter, Msg, False)
        Case "Cancel", "Undo", "Redo"
            Exit Function
    End Select
    Set SelectElement = Sel.Item(1)
    Sel.Clear
End Function

' 获取内部名称
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(ByVal aoj As AnyObject)
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
    ByVal T As String) _
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
    If TypeName(anyObj) = T Then
        Set GetParent_Of_T = anyObj
    Else
        Set GetParent_Of_T = GetParent_Of_T(anyObj.Parent, T)
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

' 检查字符串中是否包含指定关键字
' 忽略大小写进行检查
Private Function ExistsKey(ByVal txt As String, ByVal Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(Key)) > 0, True, False)
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
Private Function IsFilterType(ByVal ary As Variant) As Boolean
    IsFilterType = False
    Dim ErrMsg$: ErrMsg = "过滤器类型无效" + vbNewLine + _
                          "需要为Variant(String)类型的数组" + vbNewLine + _
                          "(具体请参考文档)"
    
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    
    IsFilterType = True
End Function

' 将字符串转换为变体数组
Private Function ToStrVriAry(ByVal s$) As Variant
    Dim ary As Variant: ary = Split(s, ",")
    Dim vriary() As Variant: ReDim vriary(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        vriary(i) = ary(i)
    Next
    ToStrVriAry = vriary
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
Function IsType_Of_T(ByVal oj As Object, ByVal T$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = T, True, False)
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
Function IsExists(ByVal path$) As Boolean
    IsExists = False
    Dim FSO As Object: Set FSO = GetFSO
    If FSO.FileExists(path) Then
        IsExists = True: Exit Function ' 文件
    ElseIf FSO.FolderExists(path) Then
        IsExists = True: Exit Function ' 文件夹
    End If
    Set FSO = Nothing
End Function

' 获取新文件名
''' @param:Path-完整路径
''' @return:新的完整路径
Function GetNewName$(ByVal oldPath$)
    Dim path As Variant
    path = SplitPathName(oldPath)
    path(2) = "." & path(2)
    Dim newPath$: newPath = path(0) + "\" + path(1)
    If Not IsExists(newPath + path(2)) Then
        GetNewName = newPath + path(2)
        Exit Function
    End If
    Dim TempName$, i&: i = 0
    Do
        i = i + 1
        TempName = newPath + "_" + CStr(i) + path(2)
        If Not IsExists(TempName) Then
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


                                       制的                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
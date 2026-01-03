VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAUtilityLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/
'/_/ clsVbaUtilityLib
'/_/ 主要处理数组、字典
'/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Option Explicit
'Win API
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Const VER = "0.0.16"
'调试模式
Private m_DebugMode As Boolean
Private Sub Class_Initialize()
    m_DebugMode = True
End Sub
Private Sub Class_Terminate()
End Sub
Property Get DebugMode() As Boolean
    DebugMode = m_DebugMode
End Property
Property Let DebugMode(ByVal Mode As Boolean)
    m_DebugMode = Mode
End Property
Function GetVer() As String
    GetVer = VER
End Function
'输出调试信息
'param:string
'return:
Sub Dump( _
        ByVal txt As String)
    If Not m_DebugMode Then Exit Sub
    Debug.Print txt
End Sub
'变量的输出调试
'param: v-变量
'param: Txt(opt): 头部注释
'return:
Sub DumpV( _
        ByVal X As Variant, _
        Optional ByVal txt As String = "")
    If Not m_DebugMode Then Exit Sub
    Debug.Print txt & DumpVariable(X)
End Sub
''' @param x As Variant
''' @return As String
Function DumpVariable( _
        ByVal X As Variant) As String
    On Error Resume Next
    Dim ty As String: ty = TypeName(X)
    Select Case ty
        Case "Boolean":     DumpVariable = X
        Case "Integer":     DumpVariable = X & "%"
        Case "Long":        DumpVariable = X & "&"
        #If Win64 Then
        Case "LongLong":    DumpVariable = X & "^"
        #End If
        Case "Single":      DumpVariable = X & "!"
        Case "Double":      DumpVariable = X & "#"
        Case "Currency":    DumpVariable = X & "@"
        Case "Byte":        DumpVariable = "CByte(" & X & ")"
        Case "Decimal":     DumpVariable = "CDec(" & X & ")"
        Case "Date"
            Dim D As String, t As String
            If Abs(X) >= 1 Then D = Month(X) & "/" & Day(X) & "/" & Year(X)
            If Not IsInt(X) Then t = Format(X, "h:nn:ss AM/PM")
            DumpVariable = "#" & VBA.Trim(D & " " & t) & "#"
        Case "String"
            If StrPtr(X) = 0 Then
                DumpVariable = "vbNullString"
            Else
                DumpVariable = """" & Replace(X, """", """""") & """"
            End If
        Case "Error"
            If IsMissing(X) Then
                DumpVariable = "Missing"
            Else
                DumpVariable = "CVErr(" & ReMatch(CStr(X), "\d+")(0) & ")"
            End If
        Case "ErrObject"
            DumpVariable = "Err " & X.Number
        Case "Empty", "Null", "Nothing", "Unknown"
            DumpVariable = ty
        Case Else
            If IsObject(X) Then
                Select Case ty
                    Case "Dictionary"
                        Dim keys As Variant
                        keys = X.keys()
                        Dim ar2 As Variant
                        ReDim ar2(X.count - 1)
                        Dim j As Long
                        For j = 0 To X.count - 1: ar2(j) = DumpVariable(keys(j)) & ":" & DumpVariable(X(keys(j))): Next
                        DumpVariable = "Dictionary(" & Join(ar2, ", ") & ")"
                    Case "Collection"
                        Dim ar3 As Variant
                        ReDim ar3(X.count - 1)
                        Dim k As Long
                        For k = 1 To X.count: ar3(k - 1) = DumpVariable(X.item(k)): Next
                        DumpVariable = "Collection(" & Join(ar3, ", ") & ")"
                    Case Else
                        DumpVariable = ToStr(X)
                End Select
            ElseIf IsArray(X) Then
                Dim rnk As Integer: rnk = ArrRank(X)
                If rnk = 1 Then
                    Dim lb As Long: lb = LBound(X)
                    Dim ub As Long: ub = UBound(X)
                    Dim ar As Variant
                    If ub - lb < 0 Then
                        ar = Array()
                    Else
                        Dim mx As Long: mx = 8 - 1
                        Dim xb As Long: xb = IIf(ub - lb < mx, ub, lb + mx)
                        ReDim ar(lb To xb)
                        Dim i As Long
                        For i = lb To xb: ar(i) = DumpVariable(X(i)): Next
                    End If
                    DumpVariable = "Array(" & Join(ar, ", ") & IIf(xb < ub, ", ...", "") & ")"
                Else
                    DumpVariable = Replace(ty, "()", "(" & String(rnk - 1, ",") & ")"
                End If
            Else
                Err.Raise 51
        End If
    End Select
    On Error GoTo 0
End Function
''' @param num As Variant(Of Numeric Or Date)
''' @return As Boolean
Function IsInt( _
        ByVal num As Variant) As Boolean
    If IsDate(num) Then num = CDbl(num)
    If Not IsNumeric(num) Then Err.Raise 13
    IsInt = num = Fix(num)
End Function
''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of String))
Private Function ReMatch( _
        ByVal expr As String, ByVal ptrnFind As String, _
        Optional ByVal iCase As Boolean = False) As Variant
    Dim ret As Variant: ret = Array()
    Dim regEx As Object:  Set regEx = CreateRegExp(ptrnFind, IIf(iCase, "i", ""))
    Dim ms As Object:     Set ms = regEx.Execute(expr)
    If ms.count < 1 Then: GoTo Ending
    Dim sms As Object:    Set sms = ms(0).SubMatches
    ReDim ret(sms.count)
    ret(0) = ms.item(0).value
    Dim i As Integer
    For i = 1 To UBound(ret): ret(i) = sms.item(i - 1): Next
Ending:
    ReMatch = ret
End Function
''' @param x As Variant
''' @return As String
Private Function ToStr(ByVal X As Variant) As String
    If IsObject(X) Then
        On Error GoTo Err438
        ToStr = X.ToStr()
        On Error GoTo 0
    ElseIf IsArray(X) Then
        ToStr = TypeName(X)
    Else
        ToStr = X
    End If
    GoTo Escape
Err438:
    Dim e As ErrObject: Set e = Err
    Select Case e.Number
        Case 438: ToStr = TypeName(X): Resume Next
        Case Else: Err.Raise e.Number, e.source, e.Description, e.HelpFile, e.HelpContext
    End Select
Escape:
End Function
''' @param arr As Variant(Of Array(Of T))
''' @return As Integer
Private Function ArrRank( _
        ByVal arr As Variant) As Integer
    If Not IsArray(arr) Then Err.Raise 13
    Dim X As Long
    Dim i As Integer: i = 0
    On Error Resume Next
    While Err.Number = 0: X = UBound(arr, IncrPre(i)): Wend
    ArrRank = i - 1
End Function
''' @param n As Variant
''' @param stepVal As Variant
''' @return As Variant
Private Function IncrPre( _
        ByRef n As Variant, _
        Optional ByVal stepVal As Variant = 1) As Variant
    n = n + stepVal: IncrPre = n
End Function
'*****系统函数*****
'获取APC
'param:
'return:Object(Of IApc)
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
'获取VBE
'param:
'return:Object(Of IApc)
Function GetVbe() As Object
    Set GetVbe = GetApc().VBE
End Function
'正在运行的vbproject的文件夹
'param:
'return:string
Function GetThisDir() As String
    Dim oApc As Object
    Set oApc = GetApc()
    Dim projFilePath As String
    projFilePath = oApc.ExecutingProject.VBProject.Filename
    GetThisDir = GetFso.GetParentFolderName(projFilePath)
End Function
'Scripting.Dictionary
'param:compareMode-Long
'return:Object(Of Dictionary)
Function InitDict( _
        Optional compareMode As Long = vbBinaryCompare) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.compareMode = compareMode
    Set InitDict = dict
End Function
'以字符数组为键，以数量为值创建字典
'param:txtArray-array(Of string)
'param:compareMode-Long
'return:Object(Of Dictionary)
Function InitDictByArrayCount( _
        ByVal txtArray As Variant, _
        Optional compareMode As Long = vbBinaryCompare) As Object
    Dim dict As Object
    Set dict = InitDict(compareMode)
    Dim key As Variant
    For Each key In txtArray
        If Not dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            Call dict.Add(key, 0)
        End If
    Next
    Set InitDictByArrayCount = dict
End Function
'是否为相同的字典 - 通过键和值进行比较，支持嵌套
'param:dict1-Object(Of Dictionary)
'param:dict2-Object(Of Dictionary
'return:boolean
Function IsEqualDict( _
        ByVal dict1 As Object, _
        ByVal dict2 As Object) As Boolean
    If dict1.count <> dict2.count Then
        IsEqualDict = False
        Exit Function
    End If
    Dim key As Variant
    For Each key In dict1.keys
        If Not dict2.Exists(key) Then
            IsEqualDict = False
            Exit Function
        End If
        If TypeName(dict1(key)) = "Dictionary" And TypeName(dict2(key)) = "Dictionary" Then
            ' 如果两个值都是字典，则递归比较它们
            If Not IsEqualDict(dict1(key), dict2(key)) Then
                IsEqualDict = False
                Exit Function
            End If
        ElseIf dict1(key) <> dict2(key) Then
            IsEqualDict = False
            Exit Function
        End If
    Next key
    IsEqualDict = True
End Function
'字典的深拷贝
'param:source-Object(Of Dictionary)
'return:Object(Of Dictionary)
Function DeepCopyDict( _
        ByVal source As Variant) As Object
    Dim key As Variant
    Dim newDict As Object
    Dim value As Variant
    Set newDict = CreateObject("Scripting.Dictionary")
    For Each key In source.keys
        value = source.item(key)
        If TypeName(value) = "Dictionary" Then
            Set value = DeepCopyDict(value)
        End If
        newDict.Add key, value
    Next key
    Set DeepCopyDict = newDict
End Function
'*****数组函数*****
'指定大小的数组创建 - 初始值仅为let
'param:count-long
'param:value
'return:Variant(Of Array)
Function InitArray( _
        ByVal count As Long, _
        Optional ByVal value As Variant = Empty) As Variant
    Dim ary() As Variant
    If count < 0 Then
        ary = Array()
    Else
        ReDim ary(count)
    End If
    If IsEmpty(value) Then
        InitArray = ary
    End If
    Dim i As Long
    For i = 0 To UBound(ary)
        ary(i) = value
    Next
    InitArray = ary
End Function
'创建连接的数组 - 不是连接而是创建新的数组
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Variant(Of Array)
Function JointArray( _
        ByVal ary1 As Variant, _
        ByVal Ary2 As Variant) As Variant
    Select Case True
        Case Not IsArray(ary1) And Not IsArray(Ary2)
            JointArray = Array() 'Empty
            Exit Function
        Case Not IsArray(ary1)
            JointArray = Ary2
            Exit Function
        Case Not IsArray(Ary2)
            JointArray = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(Ary2) > -1
            JointArray = Ary2
            Exit Function
        Case UBound(ary1) > -1 And UBound(Ary2) < 0
            JointArray = ary1
            Exit Function
        Case UBound(ary1) < 0 And UBound(Ary2) < 0
            JointArray = Array() 'Empty
            Exit Function
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
        JointArray = ary1
End Function
'切片
'param:ary-Variant(Of Array)
'param:startIdx-Long
'param:endIdx-Long
'return:Variant(Of Array)
Function Slice( _
        ByVal ary As Variant, _
        ByVal startIdx As Long, _
        ByVal endIdx As Long) As Variant
    If Not IsArray(ary) Then Exit Function
    If endIdx - startIdx < 0 Then Exit Function
    If startIdx < 0 Then Exit Function
    If endIdx > UBound(ary) Then Exit Function
    Dim rngAry() As Variant
    ReDim rngAry(endIdx - startIdx)
    Dim i As Long
    For i = startIdx To endIdx
        rngAry(i - startIdx) = ary(i)
    Next
    Slice = rngAry
End Function
'数组的深拷贝
'param:ary-Variant(Of Array)
'return:Variant(Of Array)
Function DeepCopyArray( _
        ByVal ary As Variant) As Variant
    If Not IsArray(ary) Then Exit Function
    DeepCopyArray = Slice(ary, 0, UBound(ary))
End Function
'数组是否一致？
'param:ary1-Variant(Of Array)
'param:ary2-Variant(Of Array)
'return:Boolean
Function IsEqualArray( _
        ByVal ary1 As Variant, _
        ByVal Ary2 As Variant) As Boolean
    IsEqualArray = False
    If Not IsArray(ary1) Or Not IsArray(Ary2) Then Exit Function
    If Not UBound(ary1) = UBound(Ary2) Then Exit Function
    Dim i As Long
    For i = 0 To UBound(ary1)
        If Not ary1(i) = Ary2(i) Then Exit Function
    Next
    IsEqualArray = True
End Function
'集合的数组化 - let/set根据第一个元素判断
'集合内不要混合Let和Set元素!!
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArray( _
        ByVal lst As Variant) As Variant
    If lst.count < 1 Then
        CollectionToArray = Array()
        Exit Function
    End If
    If IsObject(lst.item(1)) Then
        CollectionToArray = CollectionToArraySet(lst)
    Else
        CollectionToArray = CollectionToArrayLet(lst)
    End If
End Function
'-集合的数组化-set
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArraySet( _
        ByVal lst As Variant) As Variant
    Dim lstCount As Long
    lstCount = lst.count
    If lstCount < 1 Then
        CollectionToArraySet = Array()
        Exit Function
    End If
    Dim ary() As Variant
    ReDim ary(lstCount - 1)
    Dim i As Long
    For i = 1 To lst.count
        Set ary(i - 1) = lst(i)
    Next
    CollectionToArraySet = ary
End Function
'-集合的数组化-let
'param:lst-collection
'return:Variant(Of Array)
Function CollectionToArrayLet( _
        ByVal lst As Variant) As Variant
    Dim lstCount As Long
    lstCount = lst.count
    If lstCount < 1 Then
        CollectionToArrayLet = Array()
        Exit Function
    End If
    Dim ary() As Variant
    ReDim ary(lstCount - 1)
    Dim i As Long
    For i = 1 To lst.count
        Let ary(i - 1) = lst(i)
    Next
    CollectionToArrayLet = ary
End Function
'数组的集合化
'param:Variant(Of Array)
'return:collection
Function ArrayToCollection( _
        ByVal ary As Variant) As collection
    Dim lst As collection
    Set lst = New collection
    If UBound(ary) < 0 Then Exit Function
    Dim v As Variant
    For Each v In ary
        lst.Add v
    Next
    Set ArrayToCollection = lst
End Function
'数组切片
'param: arr-Variant(Of Array)
'param: StartIndex(opt)-起始索引
'param: EndIndex(opt)-结束索引
'return:- 切片后的数组
Function ArraySlice( _
        ByVal arr As Variant, _
        Optional ByVal StartIndex As Long = 0, _
        Optional ByVal EndIndex As Long = -1) As Variant
    ' EndIndex 省略时，设置为数组的最后一个索引
    If EndIndex = -1 Then
        EndIndex = UBound(arr)
    End If
    ' 错误检查
    If StartIndex < LBound(arr) Or EndIndex > UBound(arr) Or StartIndex > EndIndex Then
        ArraySlice = Array() ' 返回空数组
        Exit Function
    End If
    ' 设置结果数组的大小
    Dim result() As Variant
    ReDim result(0 To EndIndex - StartIndex)
    ' 数组切片
    Dim j As Long
    j = 0
    Dim i As Long
    For i = StartIndex To EndIndex
        If IsObject(arr(0)) Then
            Set result(j) = arr(i)
        Else
            Let result(j) = arr(i)
        End If
        j = j + 1
    Next i
    ArraySlice = result
End Function
'************************************************************************
'*  冒泡排序
'*  param: arrTmp-值类型的数组
'*  return: 排序后的数组
'************************************************************************
Function BubbleSort( _
        ByVal arrTmp As Variant) As Variant
    Dim i As Long, j As Long
    Dim vTmp As Variant
    For i = LBound(arrTmp) To UBound(arrTmp) - 1
        For j = i + 1 To UBound(arrTmp)
            If ExtractNumericPart(arrTmp(j)) < ExtractNumericPart(arrTmp(i)) Then
                vTmp = arrTmp(i)
                arrTmp(i) = arrTmp(j)
                arrTmp(j) = vTmp
            End If
        Next j
    Next i
    BubbleSort = arrTmp
End Function
'------------------------------------------------------------------------
'-  提取数值部分
'-  param: sTxt-目标字符串
'-  return: 数值位置索引
'------------------------------------------------------------------------
Private Function ExtractNumericPart( _
        ByVal sTxt As String) As Long
    Dim sNumericPart As String
    Dim i As Long
    For i = 1 To Len(sTxt)
        If IsNumeric(Mid(sTxt, i, 1)) Then
            sNumericPart = sNumericPart & Mid(sTxt, i, 1)
        End If
    Next i
    ExtractNumericPart = CLng(sNumericPart)
End Function
'*****IO函数*****
'FileSystemObject
'return:Object(Of FileSystemObject)
Function GetFso() As Object
    Set GetFso = CreateObject("Scripting.FileSystemObject")
End Function
'路径/文件名/扩展名 分割
'param:FullPath-文件路径
'return:Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
Function SplitPathName( _
        ByVal fullpath As String) As Variant
    Dim path(2) As Variant ' String
    With GetFso
        path(0) = .GetParentFolderName(fullpath)
        path(1) = .GetBaseName(fullpath)
        path(2) = .GetExtensionName(fullpath)
    End With
    SplitPathName = path
End Function
'路径/文件名/扩展名 连接
'param:Path-Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
'return:文件路径
Function JointPathName( _
        ByVal path As Variant) As String
    JointPathName = ""
    If Not IsArray(path) Then Exit Function '未支持
    If Not UBound(path) = 2 Then Exit Function '未支持
    JointPathName = GetFso().BuildPath(path(0), path(1) & "." & path(2))
End Function
'获取不重复的路径
'重复时采用 (文件基本名)_(编号) 格式
'param:Path-文件路径
'return:无重复路径
Function GetUniquePath( _
        ByVal oldPath As String) As String
    Dim path() As Variant
    path = SplitPathName(oldPath)
    path(2) = "." & path(2)
    '替换文件名中不能使用的字符
    path(1) = ReplaceInvalidChars(path(1))
    Dim fso As Object
    Set fso = GetFso()
    Dim newPath As String
    newPath = path(0) + "\" + path(1)
    If Not fso.FileExists(newPath + path(2)) Then
        GetUniquePath = newPath + path(2)
        Exit Function
    End If
    Dim tempName As String
    Dim i As Long
    i = 0
    Do
        i = i + 1
        tempName = newPath + "_" + CStr(i) + path(2)
        If Not fso.FileExists(tempName) Then
            GetUniquePath = tempName
            Exit Function
        End If
    Loop
End Function
'************************************************************************
'*  将文件名中不能使用的字符替换为[_]
'*  param: sPath-文件名字符串
'*  return: 可以作为文件名使用的字符串
'************************************************************************
Function ReplaceInvalidChars( _
        ByVal sPath As String) As String
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    oRegex.Pattern = "[\/:*?""<>|]"
    oRegex.Global = True
    ReplaceInvalidChars = oRegex.Replace(sPath, "_")
End Function
'文件读取
'因为是按行读取所以需要用换行符分割
'param:path-string
'return:string
Function ReadFile( _
        ByVal path As String) As String
    ReadFile = ""
    On Error Resume Next
    With GetFso
        ReadFile = .OpenTextFile(path).ReadAll
        .Close
    End With
    On Error GoTo 0
End Function
'文件写入
'param:path-string
'param:txt-string
'param:IsAppending(opt)-追加时为True
Sub WriteFile( _
        ByVal path As String, _
        ByVal txt As String, _
        Optional ByVal IsAppending = False)
    Dim iIoMode As Integer
    iIoMode = IIf(IsAppending, 8, 2)
    With GetFso().OpenTextFile(path, iIoMode, True)
        On Error Resume Next
        .Write txt
        If Not Err.Number = 0 Then
            .Write "*** Write Errar ***"
        End If
        On Error GoTo 0
        .Close
    End With
End Sub
'睡眠
'param: dMillisecond-毫秒
Sub Wait( _
        ByVal dMillisecond As Double)
    Sleep dMillisecond
End Sub


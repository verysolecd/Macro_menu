Attribute VB_Name = "KCL"
'Attribute VB_Name = "KCL"
'vba Kantoku_CATVBA_Library ver0.1.0
'KCL.bas - �Զ���VBA��
Option Explicit

Private mSW& ' ���ʼʱ��

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

' ��������� - ѭ��ѡ����Ŀ
Sub CATMain()
    Dim msg$: msg = "��ѡ����Ŀ : ��ESC���˳�"
    Dim SI As AnyObject
    Dim doc As Document: Set doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(msg)
        If IsNothing(SI) Then Exit Do
        Stop
    Loop
End Sub

'*****CATIA��غ���*****
' ����Ƿ����ִ�в���
''' @param:DocTypes-array(string),string ָ����ִ�в������ĵ�����
''' @return:Boolean
Function CanExecute(ByVal docTypes As Variant) As Boolean
    CanExecute = False
    
    If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Function
    End If
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",") '������ת����
    
    If Not checkFilterType(docTypes) Then Exit Function '��������飬���������˳�
    
    Dim ErrMsg As String
    ErrMsg = "��֧�ֵ�ǰ��ĵ����͡�" + vbNewLine + "(" + Join(docTypes, ",") + " ���ͳ���)"
 
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
'    If UBound(filter(docTypes, TypeName(ActDoc))) < 0 Then �˴�filter������VBA�бȽϺ񷵻�����ĺ���
'        MsgBox ErrMsg, vbExclamation + vbOKOnly
'        Exit Function
'    End If
    
    CanExecute = checkDocType(docTypes)
    
    If Not CanExecute Then MsgBox ErrMsg, vbExclamation + vbOKOnly
    
End Function
Function checkDocType(ByVal docTypes As Variant)
    checkDocType = False
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",") '������ת����
    If Not checkFilterType(docTypes) Then Exit Function '��������飬���������˳�
    Dim ActDoc As Document
    On Error Resume Next
        Set ActDoc = CATIA.ActiveDocument
    On Error GoTo 0
    If ActDoc Is Nothing Then
        MsgBox "�޴򿪵��ĵ�"
        Exit Function
    End If
     If UBound(filter(docTypes, TypeName(ActDoc))) < 0 Then '�˴�filter������VBA�бȽϺ񷵻�����ĺ���
        Exit Function
    End If
 checkDocType = True
End Function




' ѡ����Ŀ
''' @param:Msg-��ʾ��Ϣ
''' @param:Filter-array(string),string ѡ�������(Ĭ��ΪAnyObject)
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

' ѡ��Ԫ��
''' @param:Msg-��ʾ��Ϣ
''' @param:Filter-array(string),string ѡ�������(Ĭ��ΪAnyObject)
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

' ��ȡ�ڲ�����
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(aoj)
    If IsNothing(aoj) Then
        GetInternalName = Empty: Exit Function
    End If
    GetInternalName = aoj.GetItem("ModelElement").InternalName
End Function

' ��ȡָ�����͵ĸ�����
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

' ��ȡBrep����
''' @param:MyBRepName-String
''' @return:String
Function GetBrepName(MyBRepName As String) As String
    MyBRepName = Replace(MyBRepName, "Selection_", "")
    MyBRepName = Left(MyBRepName, InStrRev(MyBRepName, "));"))
    MyBRepName = MyBRepName + ");WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    GetBrepName = MyBRepName
End Function

' ��ȡ����
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
            ' Ӣ��-Select an object or a command
            GetLanguage = "en"
        Case ExistsKey(st, "objet")
            ' ����-ѡ��һ�����������
            GetLanguage = "fr"
        Case ExistsKey(st, "Objekt")
            ' ����-ѡ��һ�����������
            GetLanguage = "de"
        Case ExistsKey(st, "oggetto")
            ' �������-ѡ��һ�����������
            GetLanguage = "it"
        Case ExistsKey(st, "����")
            ' ����-ѡ��һ����������
            GetLanguage = "ja"
        Case ExistsKey(st, "��ҧ�֧ܧ�")
            ' ����-ѡ��һ�����������
            GetLanguage = "ru"
        Case ExistsKey(st, "����")
            ' ����-ѡ��һ�����������
            GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13
                    ' ����-???? ?? ?? ??@unicode����ʾ��
                    GetLanguage = "ko"
                Case 23
                    ' ����-���ﳤ��ʾʾ��
                    GetLanguage = "ja"
                Case Else
                    ' �������
            End Select
    End Select
End Function
' ����Ƿ�Ϊ�ַ�������
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    IsStringAry = True
End Function

' �������������Ƿ���Ч
Private Function checkFilterType(ByVal ary As Variant) As Boolean
    checkFilterType = False
    Dim ErrMsg$: ErrMsg = "������������Ч" + vbNewLine + _
                          "��ҪΪVariant(String)���͵�����" + vbNewLine + _
                          "(������ο��ĵ�)"
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    
    checkFilterType = True
    
End Function

' ���ַ���ת��Ϊ��������
Private Function strToAry(ByVal S$) As Variant
    Dim ary As Variant: ary = Split(S, ",")
    
    Dim oAry() As Variant: ReDim oAry(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        oAry(i) = ary(i)
    Next
    
    strToAry = oAry
    
End Function

'*****ͨ����غ���*****
' �������Ƿ�ΪNothing
''' @param:OJ-Variant(Of Object)
''' @return:Boolean
Function IsNothing(ByVal oj As Variant) As Boolean
    IsNothing = oj Is Nothing
End Function

' ����Scripting.Dictionary����
''' @param:CompareMode-Long
''' @return:Object(Of Dictionary)
Function InitDic(Optional CompareMode As Long = vbBinaryCompare) As Object
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.CompareMode = CompareMode
    Set InitDic = Dic
End Function

' ����ArrayList����
''' @return:Object(Of ArrayList)Public
Function InitLst() As Object
    Set InitLst = CreateObject("System.Collections.ArrayList")
End Function

' �������Ƿ�Ϊָ������
''' @param:OJ-Object
''' @param:T-String
''' @return:Boolean
Function IsType_Of_T(ByVal oj As Object, ByVal t$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = t, True, False)
'    MsgBox TypeName(oj)
End Function


'*****������غ���*****
' �ϲ���������
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

' ��ȡ����ָ����Χ��Ԫ��
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

' ��¡����
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function CloneAry(ByVal ary As Variant) As Variant
    If Not IsArray(ary) Then Exit Function
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
End Function

' ������������Ƿ����
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


'*****IO��غ���*****
' ��ȡFileSystemObject����
''' @return:Object(Of FileSystemObject)
Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

' �ָ�·����
''' @param:FullPath-����·��
''' @return:Variant(Of Array(Of String)) (0-·�� 1-�ļ��� 2-��չ��)
Function SplitPathName(ByVal FullPath$) As Variant
    Dim path(2) As String
    With GetFSO
        path(0) = .GetParentFolderName(FullPath)
        path(1) = .GetBaseName(FullPath)
        path(2) = .GetExtensionName(FullPath)
    End With
    SplitPathName = path
End Function

' �ϲ�·����
''' @param:Path-Variant(Of Array(Of String)) (0-·�� 1-�ļ��� 2-��չ��)
''' @return:����·��
Function JoinPathName$(ByVal path As Variant)
    If Not IsArray(path) Then Stop ' �������
    If Not UBound(path) = 2 Then Stop ' �������
    JoinPathName = path(0) + "\" + path(1) + "." + path(2)
End Function

' ���·���Ƿ����
''' @param:Path-·��
''' @return:Boolean
Function isExists(ByVal path$) As Boolean
    isExists = False
    Dim fso As Object: Set fso = GetFSO
    If fso.FileExists(path) Then
        isExists = True: Exit Function ' �ļ�
    ElseIf fso.FolderExists(path) Then
        isExists = True: Exit Function ' �ļ���
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
' ��ȡ���ļ���
''' @param:Path-����·��
''' @return:�µ�����·��
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

' д���ļ�
''' @param:Path-����·��
''' @param:Txt-String
Sub WriteFile(ByVal path$, ByVal txt) '$)
    On Error Resume Next
        Call GetFSO.OpenTextFile(path, 2, True).Write(txt)
    On Error GoTo 0
End Sub

' ��ȡ�ļ�
''' @param:Path-����·��
''' @return:Variant(Of Array(Of String))
Function ReadFile(ByVal path$) As Variant
    On Error Resume Next
    With GetFSO.GetFile(path).OpenAsTextStream
        ReadFile = Split(.ReadAll, vbNewLine)
        .Close
    End With
    On Error GoTo 0
End Function


'*****��ʱ��غ���*****
' �������
Sub SW_Start()
    mSW = timeGetTime
End Sub

' ��ȡ��ʱʱ��
''' @return:Double(Unit:s)
Function SW_GetTime#()
    SW_GetTime = IIf(mSW = 0, -1, (timeGetTime - mSW) * 0.001)
End Function

Public Function GetInput(msg) As String
    Dim UserInput As String
    UserInput = InputBox(msg, "������ʾ")
    ' ����û�û���������ȡ�����򷵻�Ĭ��ֵ"XX"
    If UserInput = "" Or UserInput = "0" Then
        GetInput = ""
    Else
        GetInput = UserInput
    End If
End Function

'@@param: oPath-·��
Public Function ofParentPath(ByVal opath$)
    Dim idx
    idx = InStrRev(opath, "\")
If idx > 0 Then
        ofParentPath = Left(opath, idx)
    Else
        ofParentPath = opath
    End If
End Function
' ����ַ������Ƿ����ָ���ؼ���
' ���Դ�Сд���м��
Public Function ExistsKey(ByVal txt As String, ByVal Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(Key)) > 0, True, False)
End Function
'@@ param:ostr-ʱ���ʽ

Public Function timestamp(Optional ByVal ostr) As String
    Dim FT As String  ' ��ʽ��������
    Select Case True
        Case ExistsKey(ostr, "i"): FT = "yymmdd.hhnn"
        Case ExistsKey(ostr, "h"): FT = "yymmdd.hh"
        Case ExistsKey(ostr, "d"): FT = "yymmdd"
        Case ExistsKey(ostr, "s"): FT = "yymmdd.hhnnss"
        Case Else: FT = "yymmdd"  ' Ĭ�ϸ�ʽ������δ��ֵ���
    End Select
    timestamp = Format(Now, FT)
End Function
Function isEngPath(ByVal path As String) As Boolean
    Dim i As Long, charCode As Long
    Dim validChars As String
    ' ���������Ӣ�ķ��ţ�����·���ָ�����
    validChars = "!@#$%^&*()-_=+[]{};:'"",.<>/?\|~\/"
    ' ����·���е�ÿ���ַ�
    For i = 1 To Len(path)
        charCode = AscW(Mid(path, i, 1))
        ' ����Ƿ�ΪӢ����ĸ��A-Z, a-z��
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            GoTo NextChar  ' ��ͬ�� Continue For
        End If
        ' ����Ƿ�Ϊ���֣�0-9��
        If charCode >= 48 And charCode <= 57 Then
            GoTo NextChar  ' ��ͬ�� Continue For
        End If
        ' ����Ƿ�Ϊ�����Ӣ�ķ���
        If InStr(validChars, Mid(path, i, 1)) > 0 Then
            GoTo NextChar  ' ��ͬ�� Continue For
        End If
        ' ��������ǣ���·�������Ƿ��ַ�
        isEngPath = False
        Exit Function
NextChar:
    Next i
    
    ' �����ַ���ͨ�����
    isEngPath = True
End Function

' �˺������ڼ�������·���Ƿ���������ַ�
' ����:
'   pathToCheck - ��Ҫ����·��
' ����ֵ:
'   Boolean ���ͣ�True ��ʾ·���������ģ�False ��ʾ������
Function isPathchn(pathToCheck) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' ����������ʽģʽ��ƥ�������ַ�
    regEx.Pattern = "[\u4e00-\u9fa5]"
    regEx.IgnoreCase = True
    regEx.Global = True
    ' ִ��ƥ�䲢���ؽ��
    isPathchn = regEx.test(pathToCheck)
    Set regEx = Nothing
End Function
'@iStr string
'����ַ������һ��"iext"֮ǰ���ַ��򷵻�ԭ�ַ�
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
'����ַ�����һ��"_"֮ǰ���ַ��򷵻�ԭ�ַ�
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

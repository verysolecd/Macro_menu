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
    Dim Msg$: Msg = "��ѡ����Ŀ : ��ESC���˳�"
    Dim SI As AnyObject
    Dim Doc As Document: Set Doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(Msg)
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
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",")
    If Not IsFilterType(docTypes) Then Exit Function
    
    Dim ErrMsg As String
    ErrMsg = "��֧�ֵ�ǰ��ĵ����͡�" + vbNewLine + "(" + Join(docTypes, ",") + " ���ͳ���)"
    
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

' ѡ����Ŀ
''' @param:Msg-��ʾ��Ϣ
''' @param:Filter-array(string),string ѡ�������(Ĭ��ΪAnyObject)
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

' ѡ��Ԫ��
''' @param:Msg-��ʾ��Ϣ
''' @param:Filter-array(string),string ѡ�������(Ĭ��ΪAnyObject)
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

' ��ȡ�ڲ�����
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(ByVal aoj As AnyObject)
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

' ����ַ������Ƿ����ָ���ؼ���
' ���Դ�Сд���м��
Private Function ExistsKey(ByVal txt As String, ByVal Key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(Key)) > 0, True, False)
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
Private Function IsFilterType(ByVal ary As Variant) As Boolean
    IsFilterType = False
    Dim ErrMsg$: ErrMsg = "������������Ч" + vbNewLine + _
                          "��ҪΪVariant(String)���͵�����" + vbNewLine + _
                          "(������ο��ĵ�)"
    
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    
    IsFilterType = True
End Function

' ���ַ���ת��Ϊ��������
Private Function ToStrVriAry(ByVal s$) As Variant
    Dim ary As Variant: ary = Split(s, ",")
    Dim vriary() As Variant: ReDim vriary(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        vriary(i) = ary(i)
    Next
    ToStrVriAry = vriary
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
Function IsType_Of_T(ByVal oj As Object, ByVal T$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = T, True, False)
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
Function IsExists(ByVal path$) As Boolean
    IsExists = False
    Dim FSO As Object: Set FSO = GetFSO
    If FSO.FileExists(path) Then
        IsExists = True: Exit Function ' �ļ�
    ElseIf FSO.FolderExists(path) Then
        IsExists = True: Exit Function ' �ļ���
    End If
    Set FSO = Nothing
End Function

' ��ȡ���ļ���
''' @param:Path-����·��
''' @return:�µ�����·��
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


                                       �Ƶ�                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
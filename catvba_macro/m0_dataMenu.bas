Attribute VB_Name = "m0_dataMenu"
'Attribute VB_Name = "Cat_Macro_Menu_Model"
' �˴������ڻ�ȡ��˵������������Ϣ��չʾ�˵�����
Const FormTitle = "Macro"
'----- �˵���������Ϣ ---------------------------------------
' �˵�����ʾ����
' True - ��ģ̬��ʾ  False - ģ̬��ʾ
Private Const MENU_SHOW_TYPE = True
' �˵�����������
' True - ���ز˵���ť  False - ��ʾ�˵���ť
Private Const MENU_HIDE_TYPE = False
' �˵������������Ϣ
' �������Ҫ�޸�
'{ ������ : ������� }
' ʾ������
Private Const groupName = _
            "{1 : ͼֽ���� }" & _
            "{2 : �����ģ }" & _
            "{3 : �ܳ�װ�� }" & _
            "{4 : ��ȡ�޸� }" & _
            "{5 : BOM����}"
'-----------------------------------------------------------------
Option Explicit
'----- ���ò��� �����޸ĳ��Ǳ�Ҫ -----------------------
' �˵�����ӳ���
Private PageMap As Object
' ��ǩӳ���
Private TagMap As Object                    ' �����ű�ǩ
Private Const TAG_S = "{"                   ' ���ÿ�ʼ��ǩ
Private Const TAG_D = ":"                   ' ���÷ָ���ǩ
Private Const TAG_E = "}"                   ' ���ý�����ǩ
Private Const TAG_GROUP = "gp"              ' �����ű�ǩ
Private Const TAG_ENTRYPNT = "ep"           ' ��ڵ��ǩ
Private Const TAG_ENTRY_DEF = "CATMain"     ' ��ڵ�Ĭ��ֵ
Private Const TAG_PJTPATH = "pjt_path"      ' ��Ŀ·����ǩ
Private Const TAG_MDLNAME = "mdl_name"      ' ģ�����Ʊ�ǩ

'-----------------------------------------------------------------
' �˵���ڵ�
Sub CATMain()
    
    ' ��ʼ���˵�����ӳ���
    Set PageMap = Get_KeyValue(groupName, True)
    
    ' ��ȡ��ť������Ϣ
    Dim ButtonInfos As Object
    Set ButtonInfos = Get_ButtonInfo()
    If ButtonInfos Is Nothing Then
        MsgBox "δ�ҵ����õĺ���Ϣ", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    ' �԰�ť��Ϣ��������
    Dim SoLst As Object
    Set SoLst = To_SortedList(ButtonInfos)
    If SoLst Is Nothing Then Exit Sub
    
    ' ��ʾ�˵�����
    Dim Menu
    Set Menu = New Cat_Macro_Menu_View
    Call Menu.Set_FormInfo(SoLst, PageMap, FormTitle, MENU_HIDE_TYPE)
    
    If MENU_SHOW_TYPE Then
        Menu.Show vbModeless
    Else
        Menu.Show vbModal
    End If
End Sub
'******* �������� *********
' ��ȡ�갴ť��������Ϣ
' ����  :
' ����ֵ: lst(Dict)
Private Function Get_ButtonInfo() As Object
    Set Get_ButtonInfo = Nothing
    
    Dim Apc As Object: Set Apc = GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    Dim PjtPath As String: PjtPath = ExecPjt.DisplayName
    
    Dim AllComps As Object
    Set AllComps = GetModuleLst(ExecPjt.ProjectItems.VBComponents)
    If AllComps Is Nothing Then Exit Function
    
    Dim Comp As Object 'VBComponent
    Dim Mdl As Object 'CodeModule
    Dim DecCode As String
    Dim DecCnt As Long
    Dim MdlInfo As Object
    Dim CanExecMethod As String
    Dim BtnInfos As Object: Set BtnInfos = KCL.InitLst()
    
    For Each Comp In AllComps
        Set Mdl = Comp.CodeModule
        ' ��ȡ��������
        DecCnt = Mdl.CountOfDeclarationLines
        If DecCnt < 1 Then GoTo Continue
        ' ��ȡ��������
        DecCode = Mdl.Lines(1, Mdl.CountOfDeclarationLines)
        ' ��ȡ������Ϣ
        Set MdlInfo = Get_KeyValue(DecCode)
        If MdlInfo Is Nothing Then GoTo Continue
        ' ��������Ϣ
        If Not MdlInfo.Exists(TAG_GROUP) Then GoTo Continue
        If IsNumeric(MdlInfo(TAG_GROUP)) Then
            MdlInfo(TAG_GROUP) = CLng(MdlInfo(TAG_GROUP))
        Else
            GoTo Continue
        End If
        Debug.Print TypeName(MdlInfo(TAG_GROUP)) & " : " & MdlInfo(TAG_GROUP)
        If Not PageMap.Exists(MdlInfo(TAG_GROUP)) Then GoTo Continue
        
        ' �����ڵ㷽��
        CanExecMethod = vbNullString
        If MdlInfo.Exists(TAG_ENTRYPNT) Then
            If Exist_Method(Mdl, MdlInfo(TAG_ENTRYPNT)) Then
                CanExecMethod = MdlInfo(TAG_ENTRYPNT)
            Else
                GoTo Try_TAG_ENTRY_DEF
            End If
        Else
Try_TAG_ENTRY_DEF:
            If Exist_Method(Mdl, TAG_ENTRY_DEF) Then
                 CanExecMethod = TAG_ENTRY_DEF
            End If
        End If
        If CanExecMethod = vbNullString Then GoTo Continue
        Set MdlInfo = Push_Dic(MdlInfo, TAG_ENTRYPNT, CanExecMethod)        
        Set MdlInfo = Push_Dic(MdlInfo, TAG_PJTPATH, PjtPath)
        Set MdlInfo = Push_Dic(MdlInfo, TAG_MDLNAME, Mdl.Name)        
        BtnInfos.Add MdlInfo
Continue:
    Next    
    If BtnInfos.Count < 1 Then Exit Function    
    Set Get_ButtonInfo = BtnInfos
End Function
' ���ֵ�����ӻ���¼�ֵ��
' ����  : Dict,vri,vri
' ����ֵ: Dict
Private Function Push_Dic(ByVal Dic As Object, _
                          ByVal Key As Variant, _
                          ByVal Item As Variant) As Object
    If Dic.Exists(Key) Then
        Dic(Key) = Item
    Else
        Dic.Add Key, Item
    End If
    Set Push_Dic = Dic
End Function
' ���ַ�������ȡ������Ϣ - ��ת��Ϊ������
' ����  : str,Opt_bool
' ����ֵ: Dict
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
    
    Dim Matches As Object
    Set Matches = Reg.Execute(txt)
    Set Reg = Nothing
    
    If Matches.Count < 1 Then Exit Function
    
    Dim Dic As Object: Set Dic = KCL.InitDic(vbTextCompare)
    Dim Match As Object, SubMatchs As Object
    Dim Key As Variant, Var As Variant
    
    For Each Match In Matches
        Set SubMatchs = Match.SubMatches
        
        If SubMatchs.Count < 2 Then GoTo Continue
        
        Key = Trim(Replace(SubMatchs(0), """", ""))
        If Len(Key) < 1 Then GoTo Continue
        If KeyToLong Then Key = CLng(Key)
        
        Var = Trim(Replace(SubMatchs(1), """", ""))
        If Len(Var) < 1 Then GoTo Continue
        
        Set Dic = Push_Dic(Dic, Key, Var)
Continue:
    Next
    
    If Dic.Count < 1 Then Exit Function
    
    Set Get_KeyValue = Dic
End Function
' ����ť��Ϣ����������
' ����  :lst(Dict)
' ����ֵ: Dict(lst(Dict))
Private Function To_SortedList(ByVal Infos As Object) As Object
    Set To_SortedList = Nothing
    
    Dim SoLst As Object
    Set SoLst = CreateObject("System.Collections.SortedList")
    Dim Lst As Object
    
    Dim Info As Object
    For Each Info In Infos
        If SoLst.ContainsKey(Info(TAG_GROUP)) = True Then
            SoLst(Info(TAG_GROUP)).Add Info
        Else
            Set Lst = KCL.InitLst()
            Lst.Add Info
            SoLst.Add Info(TAG_GROUP), Lst
        End If
    Next
    
    If SoLst.Count < 1 Then Exit Function
    
    ' ��ģ����������
    Dim i As Long
    Dim InfoDic As Object: Set InfoDic = KCL.InitDic(vbTextCompare)
    For i = 0 To SoLst.Count - 1
        InfoDic.Add SoLst.GetKey(i), Sort_by(SoLst.GetByIndex(i))
    Next
    
    Set To_SortedList = InfoDic
End Function
' ��ģ����������
' ����  :lst(Dict)
' ����ֵ: lst(Dict)
Private Function Sort_by(ByVal Lst As Object) As Object
    Dim tmp As Object
    Dim i As Long, j As Long
    Set tmp = Lst(0)
    For i = 0 To Lst.Count - 1
        For j = Lst.Count - 1 To i Step -1
            If Lst(i)(TAG_MDLNAME) > Lst(j)(TAG_MDLNAME) Then
                Set tmp = Lst(i)
                Set Lst(i) = Lst(j)
                Set Lst(j) = tmp
            End If
        Next j
    Next i
    Set Sort_by = Lst
End Function
'******* APC/VBE *********
' ��ȡAPC����
' ����  :
' ����ֵ: obj-IApc
Private Function GetApc() As Object
    Set GetApc = Nothing
    
    ' ��ȡVBA�汾��Ӧ��COM��������
    Dim COMObjectName$
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        MsgBox "��֧�ֵ�ǰVBA�汾", vbExclamation + vbOKOnly
        Exit Function
    #End If
    
    ' ��ȡAPC����
    Dim Apc As Object: Set Apc = Nothing
    On Error Resume Next
        Set Apc = CreateObject(COMObjectName)
    On Error GoTo 0
    
    If Apc Is Nothing Then
        MsgBox "�޷���ȡMSAPC.Apc����", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    Set GetApc = Apc
End Function
' ������ģ�����Ƿ����ָ������ - �����˽�з���
' ����  : obj-CodeModule,str
' ����ֵ: Boolean
Private Function Exist_Method(ByVal CodeMdl As Object, _
                              ByVal Name As String) As Boolean
    Dim tmp As Long
    On Error Resume Next
        tmp = CodeMdl.ProcBodyLine(Name, 0)
    On Error GoTo 0
    Exist_Method = tmp > 0
    Err.Number = 0
End Function
' ��ȡ��׼ģ���б�
' ����  : obj-VBComponents
' ����ֵ: lst(obj-VBComponent)
' vbext_ComponentType
' 1-vbext_ct_StdModule 2-vbext_ct_ClassModule 3-vbext_ct_MSForm
Private Function GetModuleLst(ByVal Itms As Object) As Object
    Set GetModuleLst = Nothing
    Dim Lst As Object: Set Lst = KCL.InitLst()
    Dim Itm As Object
    For Each Itm In Itms
        If Not Itm.Type = 1 Then GoTo Continue 'vbext_ComponentType
        Lst.Add Itm
Continue:
    Next
    If Lst.Count < 1 Then Exit Function
    Set GetModuleLst = Lst
End Function


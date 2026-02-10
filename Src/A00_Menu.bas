Attribute VB_Name = "A00_Menu"

Option Explicit

' --- Configuration ---
Const formTitle = "键盘造车手"
Private Const MENU_HIDE_TYPE = True
Private Const Menu_Modeless = True

' --- Group Definitions ---
Private Const GroupName = _
            "{1 : R&W }" & _
            "{3 : ASM }" & _
            "{4 : MDL }" & _
            "{5 : DRW }" & _
            "{7: CATIA }" & _
            "{6 : OTRS}"

Private PageMap As Object

' --- Entry Point ---
Private Const mdlname As String = "A00_Menu"
Public Sub CATMain()
    Set PageMap = get_Tagcfg(GroupName, True)
    Dim MenuItems As Object
    Set MenuItems = GetMenuItems()
    If MenuItems Is Nothing Then
        MsgBox "未找到可用的宏信息", vbExclamation
        Exit Sub
    End If

    ' 3. Sort and Organize (Adapting to View's expected format)
    Dim SoLst As Object
    Set SoLst = OrganizeForView(MenuItems)
    
    If SoLst Is Nothing Then Exit Sub

    ' 4. Show Menu
    Dim Menu As Cat_Macro_Menu_View ' Use existing View class
    Set Menu = New Cat_Macro_Menu_View
    Call Menu.Set_FormInfo(SoLst, PageMap, formTitle, MENU_HIDE_TYPE)
    
    If Menu_Modeless Then
        Menu.Show vbModeless
    Else
        Menu.Show vbModal
    End If
End Sub

' --- Core Logic: Scanning ---

' Scans the project for valid macros and returns a Collection of cls_menuCAT objects
Private Function GetMenuItems() As Collection
    
    Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    Dim pjtPath As String: pjtPath = ExecPjt.DisplayName
    
    ' Filter for Standard Modules only (Type=1)
    Dim comps As Object: Set comps = ExecPjt.ProjectItems.VBComponents
    Dim comp As Object
    
    Dim Result As New Collection
    For Each comp In comps
        If comp.Type = 1 Then ' vbext_ct_StdModule
            ProcessModule comp, pjtPath, Result
        End If
    Next
    
    If Result.count > 0 Then Set GetMenuItems = Result Else Set GetMenuItems = Nothing
End Function

' Processes a single module: parses tags and checks entry point
Private Sub ProcessModule(ByVal comp As Object, ByVal pjtPath As String, ByRef colls As Collection)
    Dim mdl As Object: Set mdl = comp.CodeModule
    If mdl.CountOfDeclarationLines < 1 Then Exit Sub
    Dim DecCode As String
    DecCode = mdl.Lines(1, mdl.CountOfDeclarationLines)
    ' 1. Parse Metadata using the Class
    Dim menuItem As New cls_menuCAT
    If Not menuItem.InitFromCode(DecCode, mdl.name, pjtPath) Then Exit Sub
    ' 2. Check if Group is valid in our PageMap
    Dim grpKey As Variant
    If IsNumeric(menuItem.GroupName) Then
        grpKey = CLng(menuItem.GroupName)
    Else
        grpKey = menuItem.GroupName
    End If
    
    If Not PageMap.Exists(grpKey) Then Exit Sub
    ' 3. Validate Entry Point Existence
    If Not MethodExists(mdl, menuItem.EntryPoint) Then
        ' Fallback: Try default CATMain if the specified EP was invalid or missing
        ' cls_menuCAT defaults to CATMain if empty, but if specified and missing, we try default?
        ' Logic from m0_dataMenu: if EP specified but missing, try TAG_ENTRY_DEF.
        If MethodExists(mdl, "CATMain") Then
            menuItem.EntryPoint = "CATMain"
        Else
            Exit Sub ' No valid entry point found
        End If
    End If
    ' 4. Add to Collection
    colls.Add menuItem
End Sub

' --- Helper Functions ---

Private Function MethodExists(ByVal mdl As Object, ByVal procName As String) As Boolean
    On Error Resume Next
    Dim line As Long
    line = mdl.ProcBodyLine(procName, 0) ' vbext_pk_Proc = 0
    MethodExists = (line > 0)
    On Error GoTo 0
End Function

' --- Adapter Logic: Object Collection -> Sorted Dictionary ---
' This bridges the gap between our new Class-based logic and the old View expecting Dictionaries
Private Function OrganizeForView(ByVal colls As Collection) As Object
    Dim SoLst As Object
    Set SoLst = CreateObject("System.Collections.SortedList")
    
    Dim item As cls_menuCAT
    Dim grpKey As Variant
    Dim subList As Object
    
    ' 1. Grouping
    For Each item In colls
        If IsNumeric(item.GroupName) Then grpKey = CLng(item.GroupName) Else grpKey = item.GroupName
        
        ' Convert Object back to Dictionary for the View
        Dim itemDict As Object
        Set itemDict = item.ToDictionary()
        
        If SoLst.ContainsKey(grpKey) Then
            SoLst(grpKey).Add itemDict
        Else
            Set subList = KCL.Initlst()
            subList.Add itemDict
            SoLst.Add grpKey, subList
        End If
    Next
    
    ' 2. Sorting (by Module Name)
    ' m0_dataMenu logic: convert SortedList values to a standard Dictionary where Key=GroupID, Value=SortedList(Dicts)
    Dim finalDic As Object: Set finalDic = KCL.InitDic(vbTextCompare)
    Dim i As Long
    Dim KEY As Variant
    Dim rawList As Object
    For i = 0 To SoLst.count - 1
        KEY = SoLst.GetKey(i)
        Set rawList = SoLst.GetByIndex(i)
        ' Sort the generic list
        finalDic.Add KEY, SortDictList(rawList)
    Next
    Set OrganizeForView = finalDic
End Function

' Simplified Bubble Sort for List of Dictionaries (Sorting by TAG_MDLNAME)
Private Function SortDictList(ByVal lst As Object) As Object
    Dim i As Long, j As Long
    Dim temp As Object
    
    For i = 0 To lst.count - 1
        For j = i + 1 To lst.count - 1
            If lst(i)("mdl_name") > lst(j)("mdl_name") Then
                Set temp = lst(i)
                Set lst(i) = lst(j)
                Set lst(j) = temp
            End If
        Next j
    Next i
    
    Set SortDictList = lst
End Function

' Identical to m0_dataMenu's parser for the Group Definitions string
Private Function get_Tagcfg(ByVal txt As String, Optional ByVal KeyToLong As Boolean = False) As Object
    Dim dic As Object: Set dic = KCL.InitDic(vbTextCompare)
    Dim Reg As Object: Set Reg = CreateObject("VBScript.RegExp")
    
    With Reg
        .Pattern = "{(.*?):(.*?)}" ''TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E'''"{(.*?):(.*?)}" ' Simplified regex based on Constants
        .Global = True
    End With
    
    Dim matches As Object: Set matches = Reg.Execute(txt)
    Dim match As Object
    Dim KEY As Variant, val As Variant
    
    For Each match In matches
        If match.SubMatches.count >= 2 Then
            KEY = Trim(match.SubMatches(0))
            val = Trim(match.SubMatches(1))
            
            If KeyToLong And IsNumeric(KEY) Then KEY = CLng(KEY)
            If dic.Exists(KEY) Then dic(KEY) = val Else dic.Add KEY, val
        End If
    Next
    
    Set get_Tagcfg = dic
End Function





VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbaModuleManegerView 
   Caption         =   "UserForm2"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970.001
   OleObjectBlob   =   "VbaModuleManegerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VbaModuleManegerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mUtil As clsVBAUtilityLib
Private mModuleMgr As clsVbaModuleManagerModel

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    Set mUtil = New clsVBAUtilityLib
    Set mModuleMgr = New clsVbaModuleManagerModel
    Me.Caption = mModuleMgr.title
    Call update_comboBox
End Sub
'** 事件 **
Private Sub btnExport_Click()
    Dim Res As Boolean
    Res = export_project
    Call update_listbox
    If Res Then
        show_msg "导出完成"
    End If
End Sub
Private Sub btnimport_Click()
    Call import_project
    show_msg "导入完成" & vbCrLf & "请保存项目"
End Sub
Private Sub btnOverwrite_Click()
    Call overwriting_project
    show_msg "导出完成"
End Sub
Private Sub btnFinish_Click()
    Call finish
End Sub
Private Sub ListBox1_Change()
    Call update_info_txt
End Sub
Private Sub ComboBox1_Change()
    Call update_listbox
End Sub
Private Sub btnOpen_Click()
    '仅打开文件夹
    Call mModuleMgr.open_folder( _
        Me.ComboBox1.ListIndex + 1 _
    )
End Sub
'** 支持 **
'结束
Private Sub finish()
    Me.Hide
    Unload Me
End Sub
'模块列表的更新
Private Sub update_listbox()
    With Me.ListBox1
        .Clear
        .ListIndex = -1
    End With
    If Me.ComboBox1.ListIndex < 0 Then Exit Sub
    Dim Name As Variant
    For Each Name In mModuleMgr.get_module_name_list(Me.ComboBox1.ListIndex + 1)
        Call Me.ListBox1.AddItem(Name)
    Next
    Dim btnEnabled As Boolean
    If mModuleMgr.has_user_data(Me.ComboBox1.ListIndex + 1) Then
        btnEnabled = True
    Else
        btnEnabled = False
    End If
    With Me
        .btnOverwrite.Enabled = btnEnabled
        .btnImport.Enabled = btnEnabled
        .btnOpen.Enabled = btnEnabled
    End With
    '模块列表的更新
    If Me.ComboBox1.value = mModuleMgr.project_name Then
        Me.btnImport.Enabled = False
    End If
End Sub
'信息文本的更新
Private Sub update_info_txt()
    If Me.ComboBox1.ListIndex < 0 Then
        Me.TextBox1.Text = vbNullString
        Exit Sub
    End If
    Dim value As String
    If Me.ListBox1.ListIndex < 0 Then
        value = ""
    Else
        value = Me.ListBox1.value
    End If
    Me.TextBox1.Text = mModuleMgr.get_module_info( _
        Me.ComboBox1.ListIndex + 1, _
        value _
    )
End Sub
'ComboBox初始设置
Private Sub update_comboBox()
    Dim projects As collection
    Set projects = mModuleMgr.get_project_name_list()
    If projects.count < 1 Then
        MsgBox "目标项目不存在"
        Call finish
        Exit Sub
    End If
    Dim i As Long
    For i = 1 To projects.count
        Call Me.ComboBox1.AddItem(projects.item(i))
    Next
    ComboBox1.ListIndex = 0
End Sub
'List中指定文字的索引获取
'param: value-搜索文字
'param: lst-搜索目标集合
'return: 对应索引
Function get_index_by_list( _
        ByVal value As Variant, _
        ByVal lst As collection) As Long
    Dim i As Long
    For i = 1 To lst.count
        If lst.item(i) = value Then
            get_index_by_list = i - 1
            Exit Function
        End If
    Next
    get_index_by_list = -1
End Function
'项目的导入
Private Sub import_project()
    Dim projIdx As Long
    projIdx = Me.ComboBox1.ListIndex + 1
    Call mModuleMgr.import_project( _
        projIdx _
    )
    Call update_listbox
End Sub
'项目的导出
Private Function export_project() As Boolean
    export_project = False
    Dim projIdx As Long
    projIdx = Me.ComboBox1.ListIndex + 1
    Dim msg As String
    msg = "在CATVBA文件的文件夹内创建吗？" & vbCrLf & _
        "(是-文件夹内创建 否-对话框指定)"
    Select Case MsgBox(msg, vbYesNoCancel + vbQuestion, mModuleMgr.title)
        Case vbYes
            '项目文件夹内
            mModuleMgr.export_project_child_folder ( _
                projIdx _
            )
        Case vbNo
            '对话框指定
            Dim dirPath As String
            dirPath = get_folder_path()
            If dirPath = vbNullString Then Exit Function
            Dim path As String
            path = mModuleMgr.get_dir_by_project_name( _
                projIdx, _
                dirPath _
            )
            Call mModuleMgr.export_project( _
                projIdx, _
                path _
            )
        Case Else
            '取消
            Exit Function
    End Select
    export_project = True
End Function
'文件夹路径获取对话框
'return: 文件夹路径
Private Function get_folder_path() As String
    Dim dirPicker As New clsFolderPicker
    get_folder_path = dirPicker.show_folder_picker()
'    Dim folderPath As String
'    With Application.FileDialog(msoFileDialogFolderPicker)
'        .title = "请选择文件夹"
'        If .Show = -1 Then
'            folderPath = .SelectedItems(1)
'            get_folder_path = folderPath
'        Else
'            get_folder_path = vbNullString
'        End If
'    End With
End Function
'项目的覆盖导出
Private Sub overwriting_project()
    Call mModuleMgr.overwriting_project( _
        Me.ComboBox1.ListIndex + 1 _
    )
End Sub
'消息
Private Sub show_msg( _
        ByVal msg As String)
    MsgBox msg, vbOKOnly, Me.Caption
End Sub


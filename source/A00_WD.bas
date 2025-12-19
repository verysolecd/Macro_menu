Attribute VB_Name = "A00_WD"
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
'标题格式为 %Title <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_path  是否导出到当前路径
' %UI CheckBox  chk_tm  是否更新时间戳到CATIA零件号？
' %UI CheckBox chk_log  本次导出是否更新日志？
' %UI TextBox   txt_log  请输入更新内容(不必输入时间)
' %UI Button btnOK  确定
' %UI Button btncancel  取消
' %Title 现在要导出stp我请问你?


Sub WD2()
    Dim Apc As Object: Set Apc = KCL.GetApc()
       Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
       On Error Resume Next
          On Error Resume Next
            Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
               Error.Clear
           On Error GoTo 0
         
   
    ' 1. 调用 UI 生成器，获取返回值（字典）
    Dim uiData As Object
    Set uiData = mdl2wd(mdl)
    
    ' 2. 检查是否点击了确定 (btnOK)
    If uiData("Status") <> "btnOK" Then
        MsgBox "用户取消了操作"
        Exit Sub
    End If
    
    ' 3. 根据返回的字典执行业务逻辑
    ' 示例：读取 chk_path
    If uiData.Exists("chk_path") And uiData("chk_path") = True Then
        MsgBox "执行功能：导出到当前路径"
        ' Call ExportToCurrentPath()
    End If
    
    ' 示例：读取 txt_log
    If uiData.Exists("txt_log") Then
        MsgBox "日志内容: " & uiData("txt_log")
    End If
    
 End Sub




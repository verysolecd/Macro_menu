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
    Set oDoc = CATIA.ActiveDocument
       Dim Apc As Object: Set Apc = KCL.GetApc()
       Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
       On Error Resume Next
          On Error Resume Next
            Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
               Error.Clear
           On Error GoTo 0
      Call mdl2wd(mdl)
 End Sub




Attribute VB_Name = "A00_WD"
'------宏信息------


'------窗体标题------
'标题格式为 %Title <Caption/Text>
' %%Title 现在要导出stp我请问你?

'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
'------控件清单-------

' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_path  是否导出到当前路径
' %UI CheckBox  chk_tm  是否更新时间戳到CATIA零件号？
' %UI CheckBox chk_log  本次导出是否更新日志？
' %UI TextBox   txt_log  请输入更新内容(不必输入时间)
' %UI Button btnOK  当前路径
' %UI Button btnsel  选择路径
' %UI Button btncancel  取消
' %UI Button btncancel  取消
' %UI CheckBox chk_3  本次导出是否更新日志？
Option Explicit

Sub WD2()
   
'   Dim oFrm: Set oFrm = New Cls_DynaFrm
   Dim frmDic: Set frmDic = getFrmDic ' oFrm.Res
   
   clName = ""
   
   '===首选按钮类执行类型
    
  If frmDic(clName) = " " Then  ' 2. 检查是否点击了确定 (btnOK)
  
     Select Case frmDic(clName)
     
     Case True:
     Case False
     
     End Select
  End If
     
    End If
    If frmDic("Status") <> "btnOK" Then   ' 2. 检查是否点击了确定 (btnOK)
        MsgBox "用户取消了操作"
        Exit Sub
    End If
    
    ' 3. 根据返回的字典执行业务逻辑

    If frmDic.Exists("chk_path") And frmDic("chk_path") = True Then      ' 示例：读取 chk_path
        MsgBox "执行功能：导出到当前路径"
        ' Call ExportToCurrentPath()
    End If
    
    ' 示例：读取 txt_log
    If frmDic.Exists("txt_log") Then
        MsgBox "日志内容: " & frmDic("txt_log")
    End If
    
 End Sub




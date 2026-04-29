Attribute VB_Name = "Module2"
Sub TestCustomPopup()
    Dim result As PopupResult
    
    ' 创建一个350x220的弹窗
    CreatePopup "自定义弹窗示例"
    
    ' 添加复选框
    AddCheckbox "chkOption1", "启用功能1", 20, 20, 120, 18, True
    AddCheckbox "chkOption2", "启用功能2", 20, 45, 120, 18, False
    
    ' 添加文本框
    AddTextBox "txtName", "姓名:", 20, 80, 60, 18, "请输入姓名"
    AddTextBox "txtAge", "年龄:", 20, 110, 60, 18, "18"
    
    ' 显示弹窗
    result = ShowPopup()
    
    ' 处理结果
    If result = prOK Then
        MsgBox "功能1: " & GetValue("chkOption1") & vbCrLf & _
               "功能2: " & GetValue("chkOption2") & vbCrLf & _
               "姓名: " & GetValue("txtName") & vbCrLf & _
               "年龄: " & GetValue("txtAge"), vbInformation, "用户输入"
    Else
        MsgBox "用户取消了操作", vbExclamation, "提示"
    End If
End Sub



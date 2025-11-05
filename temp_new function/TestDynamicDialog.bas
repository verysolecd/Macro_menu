Attribute VB_Name = "TestDynamicDialog"
Option Explicit

' 这是一个测试宏，用于演示如何使用新的动态窗体

Public Sub ShowMyNewDialog()
    ' 1. 创建窗体实例
    Dim promptForm As New frmCustomDialog
    
    ' 2. 使用 With 语句来方便地配置窗体
    With promptForm
        .caption = "通用动态对话框"
        
        ' 3. 动态添加各种控件
        .AddLabel "lblInfo", "这是一个动态生成的标签，可以点击。", 10, 10
        .AddCheckBox "chkOption1", "启用选项1", 10, 40
        .AddTextBox "txtUserInput", "在这里输入...", 10, 70, 200, 20
        .AddButton "btnOK", "确定", 10, 110, 80, 25
        .AddButton "btnCancel", "取消", 100, 110, 80, 25
        
        ' 4. 调整窗体大小以适应内容
        .AdjustSize
        
        ' 5. 显示窗体 (vbModal表示它会阻止代码继续执行，直到窗体关闭)
        .Show vbModal
    End With
    
    ' 窗体关闭后，对象会被自动清理
End Sub

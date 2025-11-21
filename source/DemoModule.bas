Attribute VB_Name = "DemoModule"
'''
' @module DemoModule
' @description 这是一个演示如何使用 DynamicFormBuilder 的示例模块。
'              通过在这个文件中编写特定格式的注释，然后运行 RunDynamicFormBuilder
'              即可动态生成一个用户窗体。
'''

' ---------------------------------------------------------------------------
' UI 声明区域
'
' 使用以下格式来声明你想要的控件：

' ' %UI <ControlType> <ControlName> <Left> <Top> <Width> <Height> <Caption/Text>
'
' - ControlType: 支持 button, label, textbox, checkbox, optionbutton
' - ControlName: 控件的唯一名称 (用于编程)
' - Left, Top, Width, Height: 控件的位置和尺寸 (单位: point)
' - Caption/Text: 控件上显示的文字
'
' ---------------------------------------------------------------------------

' %UI Label lblTitle 10 10 260 20 这是一个动态生成的窗体
' %UI Label lblName 10 40 80 20 名称:
' %UI TextBox txtName 90 38 180 22 请输入...
' %UI CheckBox chkEnable 10 70 150 20 启用高级选项
' %UI Button btnOK 110 110 80 25 确定
' %UI Button btnCancel 200 110 80 25 取消


'''
' @description
' 运行此过程来触发动态窗体的生成。
' 它会调用 DynamicFormBuilder 模块中的主函数，并把当前模块的名称传递过去。
'''
Option Explicit

Public Sub RunDynamicFormBuilder()
    ' 调用生成器，并传入当前模块的名称 "DemoModule"
    CreateFormFromModuleDeclarations "DemoModule"
End Sub


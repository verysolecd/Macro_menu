VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} wd 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "wd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "wd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 模块：modStyle（简化版）
' 布局常量（核心简化点）
Private Const FORM_WIDTH As Integer = 800 ' 窗体固定宽度
Private Const LEFT_MARGIN As Integer = 15 ' 所有控件左对齐的左边距
Private Const CONTROL_SPACING As Integer = 10 ' 控件间垂直间距
Private Const TOP_START As Integer = 15 ' 第一个控件的顶部起始位置
' 控件默认尺寸
Private Const LABEL_HEIGHT As Integer = 15 ' 标签高度
Private Const BTN_WIDTH As Integer = 80 ' 按钮宽度
Private Const BTN_HEIGHT As Integer = 25 ' 按钮高度
Private Const INPUT_WIDTH As Integer = 250 ' 输入框宽度（=窗体宽-2*左边距）
Private Const INPUT_HEIGHT_SINGLE As Integer = 20 ' 单行输入框高度
Private Const INPUT_HEIGHT_MULTI As Integer = 60 ' 多行输入框高度
Private Const OPTION_HEIGHT As Integer = 18 ' 单选/复选框高度
' 样式常量（保持美观）
Private Const FONT_NAME As String = "微软雅黑"
Private Const FONT_SIZE As Integer = 10
Private Const FORM_BACKCOLOR As Long = &H8000000F ' 浅灰背景
Private Const BTN_BACKCOLOR As Long = &H8000000D ' 按钮灰蓝


Private Sub UserForm_Click()

End Sub

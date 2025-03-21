Attribute VB_Name = "global_var"
Public gPrd As Object
Public rootPrd  As Object  '全局产品obj
Public startrow, lastrow  '全局excel行定义
Public xlApp As Object  '全局excelcom组件
Public gwb As Object
Public gws  As Object
Public pdm, xlm  '全局类实例
Public allPN  '全局遍历dict

Sub clearall()


MsgBox gPrd.Value
End Sub

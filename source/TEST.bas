Attribute VB_Name = "TEST"
' === 如何使用 ===
' 创建一个模块并调用 ActivateOrOpenExplorer
Sub TestMyCode()
'    ' 请将下面的路径替换为您想测试的实际文件夹路径
'
'  Dim oPath
'     oPath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
'
'    Debug.Print "正在尝试激活或打开: " & oPath
'    KCL.openpath oPath


    str1 = "ProductDocument,PartDocument"
    MsgBox LCase(str1)



End Sub


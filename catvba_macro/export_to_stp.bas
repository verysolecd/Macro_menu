Attribute VB_Name = "export_to_stp"
' 导出当前活动文档为STP文件
Option Explicit

Sub ExportToSTP()
    ' 检查是否可以执行操作
    If Not KCL.CanExecute("PartDocument,ProductDocument") Then Exit Sub
    
    Dim doc As Document
    Set doc = CATIA.ActiveDocument
    
    ' 让用户选择保存路径和文件名
    Dim filePath As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "请选择保存STP文件的位置"
        .InitialFileName = "example.stp"
        .Filters.Clear
        .Filters.Add "STEP 文件", "*.stp"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            If Right(filePath, 4) <> ".stp" Then
                filePath = filePath & ".stp"
            End If
        Else
            MsgBox "未选择保存路径，操作取消。", vbExclamation
            Exit Sub
        End If
    End With
    
    If filePath = "" Then
        MsgBox "未输入有效的保存路径，操作取消。", vbExclamation
        Exit Sub
    End If
    
    ' 导出为STP文件
    On Error Resume Next
    doc.ExportData filePath, "stp"
    If Err.Number <> 0 Then
        MsgBox "导出失败：" & Err.Description, vbCritical
    Else
        MsgBox "文件已成功导出到：" & filePath, vbInformation
    End If
    On Error GoTo 0
End Sub
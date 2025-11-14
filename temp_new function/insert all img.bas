Sub InsertPicturesToColumnC()

    ' --- 用户配置 ---
    ' 请将 "Sheet1" 替换为您的工作表名称
    Const wsName As String = "Sheet1"
    ' 请将 "C:\Your\Image\Folder\" 替换为您存放图片的文件夹路径 (注意最后的反斜杠)
    Const imageFolderPath As String = "C:\Your\Image\Folder\"
    ' 假设图片是 .jpg 格式，如果不是，请修改为 .png, .bmp 等
    Const imageExtension As String = ".jpg"
    ' --- 配置结束 ---

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim partNumber As String
    Dim imagePath As String
    Dim targetCell As Range
    Dim pic As Object ' Shape

    ' 获取工作表对象
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "错误：找不到名为 '" & wsName & "' 的工作表。", vbCritical
        Exit Sub
    End If

    ' 找到B列的最后一行数据
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' 从第二行开始循环 (假设第一行是标题)
    For i = 2 To lastRow
        ' 获取B列的零件号
        partNumber = ws.Cells(i, "B").Value

        ' 检查零件号是否为空
        If Trim(partNumber) <> "" Then
            ' 构建完整的图片路径
            imagePath = imageFolderPath & partNumber & imageExtension

            ' 检查图片文件是否存在
            If Dir(imagePath) <> "" Then
                ' 获取C列的目标单元格
                Set targetCell = ws.Cells(i, "C")

                ' (可选) 清除单元格中可能存在的旧图片，防止重复插入
                Dim oldPic As Shape
                For Each oldPic In ws.Shapes
                    If oldPic.Type = msoPicture And oldPic.TopLeftCell.Address = targetCell.Address Then
                        oldPic.Delete
                    End If
                Next oldPic

                ' 插入图片并将其位置和大小调整为与目标单元格匹配
                ' 参数: Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height
                Set pic = ws.Shapes.AddPicture(imagePath, msoFalse, msoTrue, targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)
                
                ' (可选) 设置图片随单元格移动和缩放
                pic.Placement = xlMoveAndSize

            Else
                ' (可选) 如果图片未找到，可以在C列单元格中写入提示信息
                ws.Cells(i, "C").Value = "图片未找到"
            End If
        End If
    Next i

    MsgBox "图片插入完成！", vbInformation

End Sub
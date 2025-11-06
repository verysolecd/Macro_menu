Sub WriteProductsToExcel_Elegant()
  
   
    startRow = 2 ' 起始行
    targetAttrs = Array(0, 2, 4) ' 需提取的属性索引（0-based）
    targetCols = Array(2, 4, 6) ' 目标列号（1-based，如B、D、F列）
   
    ' 2. 数据源（替换为实际数据，示例为10000条模拟数据）
    totalProducts = 10000
    ReDim products(0 To totalProducts - 1)
    For i = 0 To totalProducts - 1
        products(i) = Array("产品" & i + 1, "色" & i Mod 5, 100 + i, "材" & i Mod 3, "2023-" & (i Mod 12 + 1) & "-" & (i Mod 28 + 1))
    Next
    ' 3. 内存数组暂存结果
    ReDim resultArr(1 To totalProducts, 1 To UBound(targetAttrs) + 1)
    For i = 1 To totalProducts
        For j = 1 To UBound(resultArr, 2)
            resultArr(i, j) = products(i - 1)(targetAttrs(j - 1))
        Next j
    Next i
    ' 4. 按列批量写入（不影响其他列）
    For j = 0 To UBound(targetCols)
        ws.Cells(startRow, targetCols(j)).Resize(totalProducts).Value = Application.Index(resultArr, , j + 1)
    Next j
    MsgBox "写入完成：" & totalProducts & "行 × " & UBound(targetCols) + 1 & "列", vbInformation
CleanUp: ' 恢复Excel设置
    With Application
        .ScreenUpdating = True: .Calculation = xlCalculationAutomatic: .EnableEvents = True
    End With
    On Error GoTo 0
End Sub
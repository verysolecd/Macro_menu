Attribute VB_Name = "BOM_Format"
'{GP:444}
'{EP:CATMain}
'{Caption:设定BOM格式}
'{ControlTipText: 按初始化模板设定BOM格式}
'{背景颜色: 12648447}


Sub CATMain()
 If Not CanExecute("ProductDocument") Then Exit Sub
     Dim opath: opath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
    Dim bfile:    bfile = opath & "\bom_recap.txt"
    Dim rootPrd: Set rootPrd = CATIA.ActiveDocument.Product
    Dim ASMConv: Set ASMConv = rootPrd.getItem("BillOfMaterial")
    Dim Ary(7) 'change number if you have more custom columns/array...
    Ary(0) = "Number"
    Ary(1) = "Part Number"
    Ary(2) = "Quantity"
    Ary(3) = "Nomenclature"
    Ary(4) = "Defintion"
    Ary(5) = "Mass"
    Ary(6) = "Density"
    Ary(7) = "Material"
'   ASMConv.SetCurrentFormat Ary
   ASMConv.SetSecondaryFormat Ary
  
    CallByName ASMConv, "Print", VbMethod, "TXT", bfile, rootPrd
   Set olns = getBomlns(bfile)
    
    bomary = Parse2ary(olns)
    
   If xlm Is Nothing Then Set xlm = New Cls_XLM
     xlm.xlshow
xlm.inject_ary bomary
    xlm.xlshow


'    Call parseStrLines


'    ' --- 3. 在 2D 图纸中创建表格 ---
'    Dim oDrwView As DrawingView
'    Set oDrwView = CATIA.ActiveDocument.Sheets.ActiveSheet.Views.item("Background View")
'
'    Dim oTable As DrawingTable
'    ' 行数 = 匹配到的数据行 + 1行表头
'    Set oTable = oDrwView.Tables.Add(20, 200, matches.count + 1, 4)
'
'    ' 设置表头
'    oTable.SetCellString 1, 1, "序号"
'    oTable.SetCellString 1, 2, "件数"
'    oTable.SetCellString 1, 3, "代号"
'    oTable.SetCellString 1, 4, "备注"
'
'    ' 填入匹配到的数据
'    Dim i As Integer
'    For i = 0 To matches.count - 1
'        Dim m As Object: Set m = matches.item(i)
'        oTable.SetCellString i + 2, 1, m.SubMatches(0) ' 序号
'        oTable.SetCellString i + 2, 2, m.SubMatches(1) ' 数量
'        oTable.SetCellString i + 2, 3, m.SubMatches(2) ' 零件号
'    Next
'
'    ' 清理
'    fso.DeleteFile sTempFile
'    MsgBox "BOM 汇总表已自动生成！"
End Sub

Function Parse2ary(lns)
    Parse2ary = Split("") ' 初始化为空数组以防万一
    If lns.count = 0 Then Exit Function
    Dim tempArr, resultArr
    tempArr = ParseLineToArray(lns(1))
    Dim colCount, i, j, maxCol
    colCount = UBound(tempArr) - LBound(tempArr) + 1
    ReDim resultArr(1 To lns.count, 1 To colCount)
    For i = 1 To lns.count
        tempArr = ParseLineToArray(lns(i))
        maxCol = UBound(tempArr) - LBound(tempArr) + 1
        If maxCol > colCount Then maxCol = colCount
        For j = 1 To maxCol
            resultArr(i, j) = tempArr(j - 1) 'tempArr 是 0-based
        Next j
        For j = maxCol + 1 To colCount
             resultArr(i, j) = ""
        Next j
    Next i
    Parse2ary = resultArr
End Function

Function getBomlns(BomTxTfile)
    Dim Fso: Set Fso = KCL.GetFso
    Dim ts: Set ts = Fso.OpenTextFile(BomTxTfile, 1)
    Dim lns: Set lns = New collection
    Dim startLn: startLn = False
    Do Until ts.AtEndOfStream
        lineContent = Trim(ts.ReadLine)
        If InStr(lineContent, "Recapitulation") > 0 Then startLn = True
    If startLn = True Then
        If Left(lineContent, 1) = "|" Then lns.Add lineContent
    End If
    Loop
    ts.Close
   Set getBomlns = lns
End Function

Function ParseLineToArray(lineStr)
    Dim tempStr:   tempStr = Trim(lineStr)
   
    If Left(tempStr, 1) = "|" Then tempStr = Mid(tempStr, 2)
    If Right(tempStr, 1) = "|" Then tempStr = Left(tempStr, Len(tempStr) - 1)

    Dim arr:    arr = Split(tempStr, "|")
    
    Dim k As Integer
    For k = LBound(arr) To UBound(arr)
        arr(k) = Trim(arr(k))
    Next k
    
    ParseLineToArray = arr
End Function
 

Attribute VB_Name = "Module2"
Sub GenerateRecapBOMToTable()
'   Dim rootPrd: Set rootPrd = CATIA.ActiveDocument.Product
'    Dim oConv: Set oConv = rootPrd.getItem("BillOfMaterial")
'
'    Dim Ary(7) 'change number if you have more custom columns/array...
'    Ary(0) = "Number"
'    Ary(1) = "Part Number"
'    Ary(2) = "Quantity"
'    Ary(3) = "Nomenclature"
'    Ary(4) = "Defintion"
'    Ary(5) = "Mass"
'    Ary(6) = "Density"
'    Ary(7) = "Material"
''    oCONv.SetCurrentFormat Ary
'    oConv.SetSecondaryFormat Ary
'

    Set doc = CATIA.ActiveDocument
    
    Set byt = doc.Sheets.ActiveSheet
    
    
    Dim ss: Set ss = CATIA.SystemService


    
    
    

    Dim sTempFile As String: sTempFile = CATIA.SystemService.Environ("TEMP") & "\bom_recap.txt"
    oConv.Print "TXT", sTempFile, oRoot

'    ' --- 2. 正则表达式解析 ---
'    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
'    Dim ts As Object: Set ts = fso.OpenTextFile(sTempFile, 1)
'    Dim sAllContent As String: sAllContent = ts.ReadAll: ts.Close
'
'    ' 我们只需要 "Recapitulation:" 之后的内容
'    Dim sRecapSection As String
'    If InStr(sAllContent, "Recapitulation") > 0 Then
'        sRecapSection = Mid(sAllContent, InStr(sAllContent, "Recapitulation"))
'    Else
'        sRecapSection = sAllContent
'    End If
'
'    ' 定义正则：匹配包含序号和零件号的行
'    ' 即使原始是 Tab 分隔，我们也可以将其预处理成竖线，或者直接正则匹配 Tab
'    ' 这里演示如何匹配你要求的竖线格式（假设你手动或通过预处理给它加了竖线）
'    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
'    With regEx
'        .Global = True
'        .MultiLine = True
'        .Pattern = "^\s*(\d+)\s*\t\s*(\d+)\s*\t\s*(.+)\s*$" ' 匹配: 序号 [Tab] 数量 [Tab] 零件号
'    End With
'
'    Dim matches As Object: Set matches = regEx.Execute(sRecapSection)
'
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

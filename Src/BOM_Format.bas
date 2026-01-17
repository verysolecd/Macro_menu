Attribute VB_Name = "BOM_Format"
'{GP:444}
'{EP:CATMain}
'{Caption:设定BOM格式}
'{ControlTipText: 按初始化模板设定BOM格式}
'{背景颜色: 12648447}
Private targetsheet
Private shts, oViews, Fct2, iformat(0 To 7), bfile


Private Sub m_init()
 
iformat(0) = "Number"
iformat(1) = "Part Number"
iformat(2) = "Quantity"
iformat(3) = "Nomenclature"
iformat(4) = "Defintion"
iformat(5) = "Mass"
iformat(6) = "Density"
iformat(7) = "Material"
 opath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
 bfile = opath & "\bom_recap.txt"

End Sub
Sub DRW_create_BomTable()
 CATIA.RefreshDisplay = False
    Call m_init
  Dim osht
  On Error Resume Next
    Set osht = Nothing
    Set osht = CATIA.ActiveDocument.sheets.item(1)
    Set osel = CATIA.ActiveDocument.Selection
  On Error GoTo 0
  If osht Is Nothing Then Exit Sub
  Set shts = osht.Parent
  Set drwDoc = shts.Parent
  Set osht = shts.ActiveSheet
  Set oViews = osht.Views
  Set bcgView = oViews.item("Background View") ' Set oView = oViews.item("Background View")
  Set mainview = oViews.item("Main View")
  Set bomview = KCL.SelectItem("请选择bom视图", "DrawingView")
  
  If bomview Is Nothing Then Exit Sub
    
    Set dprd = bomview.GenerativeBehavior.Document 'DrawingViewGenerativeBehavior/DrawingViewGenerativeBehavior
  
   ary = getPrd_BomAry(dprd, iformat)

   
 tolrow = UBound(ary, 1)
 tolcol = UBound(ary, 2)
 pos_x = 90
    pos_y = 150
On Error Resume Next


    
    Set otable = bomview.Tables.item("bbom")
    If Not otable Is Nothing Then
        pos_x = otable.X - 18
        pos_y = otable.Y - 30
            osel.Clear
            osel.Add bomview.Tables.item("bbom")
            osel.Delete
            osel.Clear
    End If
Err.Clear
On Error GoTo 0

    Set otable = bomview.Tables.Add(pos_x, pos_y, tolrow, tolcol, 10, 20)
    otable.Name = "bbom"


For i = 1 To tolrow
    
    For j = 1 To tolcol
            ostr = Trim(CStr((ary(i, j))))
    
    Call otable.SetCellString(i, j, ostr)
    
    Next j
Next i
        
  CATIA.RefreshDisplay = True



End Sub



Function getPrd_BomAry(iprd, ary)

Dim ASMConv
Set ASMConv = iprd.getItem("BillOfMaterial")
'   ASMConv.SetCurrentFormat Ary
ASMConv.SetSecondaryFormat ary
CallByName ASMConv, "Print", VbMethod, "TXT", bfile, iprd
Set olns = getBomlns(bfile)
bomary = Parse2ary(olns)
getPrd_BomAry = bomary
End Function






Sub CATMain()

 If Not CanExecute("ProductDocument") Then Exit Sub
     Dim opath: opath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
    Dim bfile:    bfile = opath & "\bom_recap.txt"
    Dim rootPrd: Set rootPrd = CATIA.ActiveDocument.Product
    Dim ASMConv: Set ASMConv = rootPrd.getItem("BillOfMaterial")
    Dim ary(7) 'change number if you have more custom columns/array...
    ary(0) = "Number"
    ary(1) = "Part Number"
    ary(2) = "Quantity"
    ary(3) = "Nomenclature"
    ary(4) = "Defintion"
    ary(5) = "Mass"
    ary(6) = "Density"
    ary(7) = "Material"
'   ASMConv.SetCurrentFormat Ary
   ASMConv.SetSecondaryFormat ary
  
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
 

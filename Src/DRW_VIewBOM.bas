Attribute VB_Name = "DRW_VIewBOM"
'{GP:5}
'{EP:DRW_create_BomTable}
'{Caption:图纸BOM}
'{ControlTipText: 在图纸中选择视图后插入产品BOM，带序号}
'{背景颜色: 12648447}

Private targetsheet
Private drwDoc, osht, shts, oViews, Fct2, iformat(0 To 7), bfile, bcgView, mainview
Private oSel
Private Const mdlname As String = "DRW_VIewBOM"
Sub DRW_create_BomTable()
 CATIA.RefreshDisplay = False
    Call m_init
On Error Resume Next

If drwDoc Is Nothing Then Exit Sub
    Set bomview = KCL.SelectItem("请选择bom视图", "DrawingView")
        If bomview Is Nothing Then Exit Sub
        
        
    Set dprd = bomview.GenerativeBehavior.Document 'DrawingViewGenerativeBehavior/DrawingViewGenerativeBehavior
        If Not IsObj_T(dprd, "Product") Then Exit Sub

tempAry = getPrd_BomAry(dprd, iformat)
 tolrow = UBound(tempAry, 1)
 tolcol = UBound(tempAry, 2)
 pos_x = 50: pos_y = 50
 Set otable = Nothing
 
 For i = 1 To bomview.Tables.count
    Set otable = bomview.Tables.item(i)
    If otbale.Name = "bbom" Then
            pos_x = otable.X - 20
            pos_y = otable.Y - 60
            bomview.Tables.Remove (i)
    End If
Next i

Err.Clear
On Error GoTo 0
    Set otable = bomview.Tables.Add(pos_x, pos_y, tolrow, tolcol, 10, 20)
    otable.Name = "bbom"
    For i = 1 To tolrow
        For j = 1 To tolcol
                ostr = Trim(CStr((tempAry(i, j))))
        Call otable.SetCellString(i, j, ostr)
        Next j
    Next i
  CATIA.RefreshDisplay = True
End Sub
Private Sub m_init()

If Not CanExecute("DrawingDocument") Then Exit Sub
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
 
  On Error Resume Next
    Set osht = Nothing
    Set osht = CATIA.ActiveDocument.sheets.item(1)
    Set oSel = CATIA.ActiveDocument.Selection
  On Error GoTo 0
  
  If osht Is Nothing Then Exit Sub
  Set shts = osht.Parent
  Set drwDoc = shts.Parent
  Set osht = shts.ActiveSheet
  Set oViews = osht.Views
  Set bcgView = oViews.item("Background View") ' Set oView = oViews.item("Background View")
  Set mainview = oViews.item("Main View")
 
End Sub
Function getPrd_BomAry(iprd, ary)
Dim ASMConv
Set ASMConv = iprd.GetItem("BillOfMaterial")
'   ASMConv.SetCurrentFormat Ary
ASMConv.SetSecondaryFormat ary
CallByName ASMConv, "Print", VbMethod, "TXT", bfile, iprd
Set olns = getBomlns(bfile)
bomary = Parse2ary(olns)
getPrd_BomAry = bomary
End Function

Sub AsmConv2xl()
     Dim opath: opath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
    Dim bfile:    bfile = opath & "\bom_recap.txt"
    Dim rootPrd: Set rootPrd = CATIA.ActiveDocument.Product
    Dim ASMConv: Set ASMConv = rootPrd.GetItem("BillOfMaterial")
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
    Dim fso: Set fso = KCL.GetFso
    Dim ts: Set ts = fso.OpenTextFile(BomTxTfile, 1)
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
 

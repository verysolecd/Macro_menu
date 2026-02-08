Attribute VB_Name = "RW_Revise"
'{GP:1}
'{Ep:EditorToolbar}
'{Caption:属性工具箱}
'{ControlTipText:打开属性管理悬浮工具条}
'  %UI Label  lblInfo  请选择操作：
'  %UI Button btnRead  读取属性
'  %UI Button btnWrite 写回属性
Option Explicit
Private prd2rv
Private Const mdlname As String = "RW_Revise"

Sub EditorToolbar()
'    If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    Dim mdlMap As Object: Set mdlMap = KCL.InitDic
    mdlMap.Add "btnRead", mdlname
    mdlMap.Add "btnWrite", mdlname
    Dim funcMap As Object: Set funcMap = KCL.InitDic
    funcMap.Add "btnRead", "readPrd"
    funcMap.Add "btnWrite", "rvme"
    If g_frm Is Nothing Then Set g_frm = KCL.newFrm(mdlname)
    g_frm.ShowToolbar mdlname, mdlMap, funcMap
End Sub

Sub readPrd()
 '---------获取待修改产品 '---------遍历修改产品及子产品
    If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = KCL.defPrd()
    Dim Prd2Read: Set Prd2Read = pdm.CurrentProduct
        If Not Prd2Read Is Nothing Then
            If gws Is Nothing Then Set xlm = New Cls_XLM
            Dim bomlns() As Bomline
            bomlns = pdm.GetReviseItems(Prd2Read)
            xlm.inject_RvData bomlns
            xlAPP.Visible = True
        End If
        Set Prd2Read = Nothing
End Sub
Sub rvme()
On Error GoTo ErrorHandler
    If xlm Is Nothing Then Err.Description = "excel错误，请检查": Exit Sub
     If pdm.CurrentProduct Is Nothing Then Exit Sub
         Dim currRow: currRow = 2
'---------遍历修改产品及子产品------
        Dim prd2rv
        Set prd2rv = pdm.CurrentProduct
            prd2rv.ApplyWorkMode (3)
        Dim children: Set children = prd2rv.Products
 
        Dim oData As Variant: oData = xlm.extract_ary
'--------map 修改ary------
    Dim iCols: iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim outputArr As Variant, tempArr(1 To 6)
    ReDim outputArr(1 To UBound(oData, 1), 1 To UBound(iCols))
    Dim i, j
    For i = 1 To UBound(outputArr, 1) '-------遍历行
  '-------遍历获取X行要修改的数据
        For j = 1 To UBound(outputArr, 2) '-------遍历该行数组后输出一个一维数组作为要的修改参数
             outputArr(i, j) = ""  '-------初始化为空数组
             If IsEmpty(oData(i, iCols(j))) = False Then outputArr(i, j) = oData(i, iCols(j))
             tempArr(j) = outputArr(i, j) '------
        Next j
    '-------遍历按行区分要修改的产品
        Select Case i
        Case 1
            Case 2  '第二行修改子总成
                Call pdm.modatt(prd2rv, tempArr)
            Case Else '其他修改子产品
                Call pdm.modatt(children.item(i - currRow), tempArr)
         End Select
    Next i
          Set prd2rv = Nothing
    MsgBox "已经修改产品"
ErrorHandler:
    If Err.Number <> 0 Then:  MsgBox Err.Description, vbCritical: Err.Clear
        Exit Sub
End Sub






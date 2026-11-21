Attribute VB_Name = "RW_attmgr"
'{GP:1}
'{Ep:AttrManager}
'{Caption:属性管理}
'{ControlTipText: 读取和写入产品属性（悬浮工具栏）}
'{BackColor:}
'  %UI Label  lblInfo  属性管理工具 #FFFF00
'  %UI Button btnRead  读取属性
'  %UI Button btnWrite 写回属性

Option Explicit
Private mWD As Cls_DynaWD
Private Const mdlname As String = "RW_attmgr"

' === 入口：显示悬浮工具栏 ===
Sub AttrManager()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    
    '==生成UItoolbar===================
    Set mWD = New Cls_DynaWD
    mWD.ShowToolbar mdlname
End Sub

' === 读取属性：从CATIA读取到Excel ===
Sub btnRead_click()
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = KCL.defPrd()
    
    Dim Prd2Read As Object
    Set Prd2Read = pdm.CurrentProduct
    
    If Not Prd2Read Is Nothing Then
        If gws Is Nothing Then Set xlm = New Cls_XLM
        Dim bomlns() As Bomline
        bomlns = pdm.GetReviseItems(Prd2Read)
        xlm.inject_RvData bomlns
        xlAPP.Visible = True
    End If
    
    Set Prd2Read = Nothing
End Sub

' === 写回属性：从Excel写回CATIA ===
Sub btnWrite_click()
    On Error GoTo ErrorHandler
    
    If xlm Is Nothing Then
        Err.Description = "excel错误，请先读取属性"
        Err.Raise vbObjectError + 1
    End If
    
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If pdm.CurrentProduct Is Nothing Then Exit Sub
    
    Dim prd2rv As Object
    Set prd2rv = pdm.CurrentProduct
    prd2rv.ApplyWorkMode (3)
    
    Dim oData As Variant
    oData = xlm.extract_ary
    
    '--- 映射列 ---
    Dim iCols As Variant
    iCols = Array(0, 2, 4, 6, 8, 10)
    
    Dim outputArr As Variant, tempArr(1 To 5)
    ReDim outputArr(1 To UBound(oData, 1), 1 To UBound(iCols))
    
    Dim i As Long, j As Long, origPN As String
    For i = 1 To UBound(outputArr, 1)
        For j = 1 To UBound(outputArr, 2)
            outputArr(i, j) = ""
            If IsEmpty(oData(i, iCols(j))) = False Then outputArr(i, j) = oData(i, iCols(j))
            tempArr(j) = outputArr(i, j)
        Next j
        
        Select Case i
            Case 1
                ' 跳过标题行
            Case 2
                ' 第二行修改子总成（根产品）
                Call pdm.modatt(prd2rv, tempArr)
            Case Else
                ' 其他修改子产品：按原始PN联动写入全部同PN实例
                origPN = oData(i, 1)
                Call pdm.modattByPN(origPN, tempArr)
        End Select
    Next i
    
    Set prd2rv = Nothing
    MsgBox "已经修改产品", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical
    Err.Clear
End Sub


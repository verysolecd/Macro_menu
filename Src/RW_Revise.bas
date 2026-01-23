Attribute VB_Name = "RW_Revise"

'{GP:3}
'{Ep:EditorToolbar}
'{Caption:属性工具箱}
'{ControlTipText:打开属性管理悬浮工具条}

'  %UI Label  lblInfo  请选择操作：
'  %UI Button btnRead  读取属性
'  %UI Button btnWrite 写回属性


Option Explicit
Private m_frm As cls_dynaFrmMLess
Private prd2rv

Sub EditorToolbar()

    If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    
    If m_frm Is Nothing Then Set m_frm = New cls_dynaFrmMLess
    
    ' 1. 定义UI配置 (模拟 cls_dynafrm 的注释风格)
    Dim config As String
    

    config =
             
    ' 2. 定义映射关系
    Dim mdlMap As Object: Set mdlMap = CreateObject("Scripting.Dictionary")
    mdlMap.Add "btnRead", "RW_Revise"
    mdlMap.Add "btnWrite", "RW_Revise"
    
    Dim funcMap As Object: Set funcMap = CreateObject("Scripting.Dictionary")
    funcMap.Add "btnRead", "readPrd"
    funcMap.Add "btnWrite", "rvme"
    
    ' 获取当前项目路径
    Dim projPath As String
    On Error Resume Next
    projPath = KCL.GetApc().ExecutingProject.VBProject.fileName
    If Err.Number <> 0 Or projPath = "" Then
         projPath = "" ' Fallback to active context
         Err.Clear
    End If
    On Error GoTo 0

    m_frm.ShowToolbar "属性管理工具箱", config, projPath, mdlMap, funcMap
    
End Sub

Sub readPrd()
  
 '---------获取待修改产品 '---------遍历修改产品及子产品
    If pdm.CurrentProduct Is Nothing Then MsgBox "请先选择产品": Exit Sub
    Dim Prd2Read: Set Prd2Read = pdm.CurrentProduct
        If Not Prd2Read Is Nothing Then
            If gws Is Nothing Then Set xlm = New Cls_XLM
            Dim currRow: currRow = 2
            g_counter = 1
            Prd2Read.ApplyWorkMode (3)
            Dim idcol, idrow
            idcol = Array(0, 1, 3, 5, 7, 9, 11, 13, 14) '' 目标列号, 0号元素不占位置
            idrow = Array(0, 1, 2, 3, 4, 5, 6, 7, 8) ' 对应的属性索引（0-based）
            Dim tmpData(): tmpData = pdm.attLv2Prd(Prd2Read)
            xlm.inject_ary tmpData, idcol, idrow
            xlm.setxlHead ("rv")
            xlm.xlshow
                xlAPP.Visible = True
        End If
        Set Prd2Read = Nothing
End Sub

Sub rvme()
  If xlm Is Nothing Then
        Set xlm = New Cls_XLM
        On Error Resume Next
        ' 尝试附加到现有 Excel 实例
        Dim tmpApp As Object: Set tmpApp = GetObject(, "Excel.Application")
         If Not tmpApp Is Nothing Then
             Set xlm.xlAPP = tmpApp
             Set xlm.ws = tmpApp.ActiveSheet ' 假设用户正在操作目标工作表
         Else
             MsgBox "未找到活动的 Excel 实例，无法更新属性。", vbCritical
             Set xlm = Nothing
             Exit Sub
         End If
         On Error GoTo 0
  End If
     If pdm Is Nothing Then Set pdm = New Cls_PDM
     If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = pdm.getiPrd()
     If pdm.CurrentProduct Is Nothing Then: MsgBox "请先选择产品，程序将退出": Exit Sub
        Dim currRow: currRow = 2
      
'---------遍历修改产品及子产品------
        Dim prd2rv
        Set prd2rv = pdm.CurrentProduct
            prd2rv.ApplyWorkMode (3)
        Dim children: Set children = prd2rv.Products
On Error GoTo ErrorHandler
        Dim odata As Variant: odata = xlm.extract_ary
'--------map 修改ary------
    Dim iCols: iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim outputArr As Variant, tempArr(1 To 6)
    ReDim outputArr(1 To UBound(odata, 1), 1 To UBound(iCols))
    
    
    
    For i = 1 To UBound(outputArr, 1) '-------遍历行
  '-------遍历获取X行要修改的数据
        For j = 1 To UBound(outputArr, 2) '-------遍历该行数组后输出一个一维数组作为要的修改参数
             outputArr(i, j) = ""  '-------初始化为空数组
             If IsEmpty(odata(i, iCols(j))) = False Then outputArr(i, j) = odata(i, iCols(j))
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
    If Err.Number <> 0 Then: Err.Clear: MsgBox "程序错误：" & Err.Description, vbCritical
        Exit Sub

End Sub





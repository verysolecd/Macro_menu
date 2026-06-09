Attribute VB_Name = "RW_read"

'{GP:1}
'{Ep:btnRead_click}
'{Caption:读取属性}
'{ControlTipText:打开属性管理悬浮工具条}
'  %UI Label  lblInfo  请选择操作：
'  %UI Button btnRead  读取属性
'  %UI Button btnWrite 写回属性
Option Explicit
Private mWD As Cls_DynaWD
Private Const mdlname As String = "RW_read"

Sub btnRead_click()
 '---------获取待修改产品 '---------遍历修改产品及子产品
   If pdm Is Nothing Then Set pdm = New Cls_PDM
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








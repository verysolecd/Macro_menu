Attribute VB_Name = "A00_globalVar"
Public Type Bomline
    level As Integer        ' 层级
    PartNumber As String    ' 件号
    Nomenclature As String  ' 英文名称
    Definition As String    ' 中文名称
    InstanceName As String  ' 实例名
    Quantity As Long        ' 数量
    Mass As Double          ' 单重
    Material As String      ' 材质
    Thickness As Double     ' 厚度
    Density As Double       ' 密度
    UserProp1 As String
    UserProp2 As String
End Type

Public Type ParamItem
    name As String
    ParamType As String
    Value As Variant
    target As Object    ' 指向 CATIA Parameter 对象
    Description As String
End Type

Public Type PropResult
    obj As Object
    Value As Variant
    IsValid As Boolean
End Type

Public rootDoc
Public rootPrd  As Object
Public xlAPP As Object
Public gwb As Object
Public gws  As Object
Public pdm As New Cls_PDM
Public xlm As New Cls_XLM
Public g_allPN As Object
Public gPic_Path
Public g_frm As cls_dynaFrm
Public g_Btn
Private Const mdlname As String = "A00_globalVar"
Sub clearall()
End Sub





Attribute VB_Name = "Add_Properties_3D"
Option Explicit
Dim Pro, Prox '属性名称，一维数组
Dim DicP, DicPx '储存属性的字典
Dim UnloadQty As Integer
Const intRowShift = 3    '表头行数
Const intColShift = 1   '偏移列数
'Const intDefaultPro = 5  '默认属性个数


Sub CATMain()


IntCATIA
IntExcel

If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If

'**************************属性名称数组赋初值,定义全局字典************************
Pro = Array("PartNumber", "Revision", "Definition", "Nomenclature", "Source", "PartDescription", "InstanceName", "ComponentDescription")
Prox = Array()

Set DicP = CreateObject("Scripting.Dictionary") '默认属性，Key 为文件名,Value 为一维属性值数组
Set DicPx = CreateObject("Scripting.Dictionary") '用户属性，Key 为文件名,Value 为一维属性值数组
'**************************属性名称数组赋初值,定义全局字典************************

On Error Resume Next
Err.Clear
Dim sExlTp As String
sExlTp = oCATVBA_Folder("Template").Path & "\" & "Properties.xls"
'Debug.Print sExlTp.Name
Set objExcel = objExcel.workbooks.Open(sExlTp)
    If Err.Number <> 0 Then
        MsgBox "无法取得Excel属性模板" & vbCrLf & sExlTp & vbCrLf & "-程序退出!"
        Exit Sub
    End If
Err.Clear

objExcel.Parent.DisplayAlerts = False
objExcel.Parent.ScreenUpdating = False

Set objExcel = objExcel.Sheets.Item("Properties")  '工作表

'####调整位置，优先显示用户自定义属性，即使为空############
Dim userdefname
userdefname = Array()
Dim x As Integer
Dim y As Integer
x = intColShift + UBound(Pro) + 2
For y = x To 100 + x
    Dim temps As String
    temps = CStr(objExcel.cells(intRowShift - 1, y).Value)
    If temps <> "" Then
    ReDim Preserve Prox(UBound(Prox) + 1)
    Prox(UBound(Prox)) = temps
    ReDim Preserve userdefname(UBound(userdefname) + 1)
    userdefname(UBound(userdefname)) = temps
    End If
    'ReDim Preserve userdefname(y - x)
    'userdefname(y - x) = CStr(objExcel.cells(intRowShift - 1, y).Value)
    'Debug.Print temps
Next
'####调整位置，优先显示用户自定义属性，即使为空############
UnloadQty = 0 '未加载文件数量
Dim oProd As Product
Set oProd = oActDoc.Product
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False

oProd.ApplyWorkMode DESIGN_MODE
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
'-----------设置设计模式-------------
GetProFromCATIA2 oProd
    If UnloadQty <> 0 Then
        MsgBox "该产品共有CATIA文件 " & DicP.Count & " 个" & vbCrLf & "该产品共(将)有用户自定义属性 " & UBound(Prox) + 1 & " 个，" & _
                vbCrLf & "其中包含了属性模板规定的用户属性" & UBound(userdefname) + 1 & "个." & vbCrLf & vbCrLf & "共有 " & UnloadQty & " 个文件未加载，未纳入统计."
    Else
        MsgBox "该产品共有CATIA文件 " & DicP.Count & " 个" & vbCrLf & "该产品共(将)有用户自定义属性 " & UBound(Prox) + 1 & " 个，" & _
                vbCrLf & "其中包含了属性模板规定的用户属性" & UBound(userdefname) + 1 & "个."
    End If
CATIA.DisplayFileAlerts = True
'####调整位置，优先显示用户自定义属性，即使为空############



objExcel.Range(objExcel.cells(intRowShift - 1, intColShift + UBound(Pro) + 2), objExcel.cells(intRowShift - 1, intColShift + UBound(Pro) + 2 + UBound(Prox))).Value = Prox
'**************** 循环填充数据******************
Dim k, j
k = Array()
k = DicP.keys()

For j = 0 To UBound(k)

objExcel.cells(intRowShift + j, intColShift) = k(j)  'CATIA文件名称栏

objExcel.Range(objExcel.cells(intRowShift + j, intColShift + 1), _
               objExcel.cells(intRowShift + j, intColShift + 1 + UBound(Pro))).Value = DicP.Item(k(j))   '默认属性值
               
objExcel.Range(objExcel.cells(intRowShift + j, intColShift + 1 + UBound(Pro) + 1), _
               objExcel.cells(intRowShift + j, intColShift + 1 + UBound(Pro) + 1 + UBound(Prox))).Value = DicPx.Item(k(j)) '用户属性值
               
Next
'####调整位置，优先显示用户自定义属性，即使为空############
'x = intColShift + 1 + UBound(Pro) + 1 + UBound(Prox)
'For y = 0 To UBound(userdefname)
'    If userdefname(y) <> "" Then
'        Dim fi As Boolean
'        fi = False
'        For j = 0 To UBound(Prox)
'            If userdefname(y) = Prox(j) Then
'                fi = True
'                Exit For
'            End If
'        Next
'        If fi = False Then
'            objExcel.cells(intRowShift - 1, x + 1) = userdefname(y)
'            x = x + 1
'        End If
'    End If
'Next
'####调整位置，优先显示用户自定义属性，即使为空############
'objExcel.cells(intRowShift - 1, intColShift).Resize(1, 1 + UBound(Pro) + 1 + UBound(Prox) + 1).AutoFilter
objExcel.Parent.Parent.DisplayAlerts = True
objExcel.Parent.Parent.ScreenUpdating = True
objExcel.Parent.Parent.Visible = True

End Sub

Sub GetProFromCATIA(oProduct) ' 输出当前Part或Product属性的字典

On Error Resume Next
Dim Docname1 As String
If Batch_ReName.ProductIsComponent(oProduct) Then
    Docname1 = oProduct.ReferenceProduct.Parent.Name & "|?|" & oProduct.PartNumber
Else
    Docname1 = oProduct.ReferenceProduct.Parent.Name
End If

If DicP.exists(Docname1) Then ' 如果这个文件的属性已取则退出，不再重复取属性
    If Err.Number <> 0 Then
        UnloadQty = UnloadQty + 1
        'MsgBox "有文件未加载，将跳过属性取值！"
    End If
    On Error GoTo 0
    Exit Sub
End If


Dim oUserPara As Object 'Parameters
Set oUserPara = oProduct.ReferenceProduct.UserRefProperties

Dim arrP   '添加默认属性的值
arrP = Array(oProduct.PartNumber, oProduct.Revision, oProduct.Definition, oProduct.Nomenclature, oProduct.Source, oProduct.DescriptionRef, oProduct.Name, oProduct.DescriptionInst)

DicP.Add Docname1, arrP


Dim arrPx    '用户属性数组
arrPx = Array()
DicPx.Add Docname1, arrPx
       If UBound(Prox) <> -1 Then       '用户属性值，数组长度与现有属性项长度相等
            Dim n
            For n = 0 To UBound(Prox)
            ReDim Preserve arrPx(n)
            arrPx(n) = ""
            Next
       End If
       
Dim i As Integer
For i = 0 To oUserPara.Count - 1       '将此零件的用户属性与已经取得的用户属性名逐个比对
    Dim ExistFlag As Boolean
    ExistFlag = False
    Dim j As Integer
                 For j = 0 To UBound(Prox)
                         If GetUserPropStr(oUserPara.Item(i + 1).Name) = Prox(j) Then '当发现这个用户属性已经在用户数组名称中存在时，修改 arrPx 数组的值
                              arrPx(j) = oUserPara.Item(i + 1).ValueAsString ',如何取得属性值？
                              ExistFlag = True
                         End If
                 Next
    If ExistFlag = False Then
           '否则用户属性名称数组增加这个名称， arrPx数组长度加一，并赋值
                Dim xx As Integer
                xx = UBound(Prox)
                ReDim Preserve Prox(xx + 1)
                Prox(xx + 1) = GetUserPropStr(oUserPara.Item(i + 1).Name)
                ReDim Preserve arrPx(xx + 1)
                arrPx(xx + 1) = oUserPara.Item(i + 1).ValueAsString
                'Debug.Print Prox(xx + 1) & " is " & arrPx(xx + 1)
                '所有已经存在的用户属性字典值数组长度也要加一，保证所有用户属性值字典的数组长度一样
                On Error Resume Next '现存数组长度为0的情况？
                Dim d  '字典的Key，文件名
                Dim m
                Dim ym '用户属性值数组
                'MsgBox IsArray(DicPx.Items())
                d = DicPx.keys()
                For m = 0 To UBound(d)
                  ym = DicPx.Item(d(m))           '用户属性数组，值的数组
                  ReDim Preserve ym(xx + 1)
                  ym(xx + 1) = ""
                  DicPx.Item(d(m)) = ym
                Next
                On Error GoTo 0
                ' This VBA Macro Developed by Charles.Tang
                ' WeChat Chtang80,CopyRight reserved
       End If
       DicPx.Item(Docname1) = arrPx
Next
End Sub
Sub GetProFromCATIA2(oProduct)
GetProFromCATIA oProduct
Dim i
If oProduct.Products.Count <> 0 Then
    For i = 1 To oProduct.Products.Count
       GetProFromCATIA2 oProduct.Products.Item(i)
    Next
End If

End Sub

Function GetUserPropStr(s As String)
GetUserPropStr = Right(s, Len(s) - InStrRev(s, "\"))
End Function


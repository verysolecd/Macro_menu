Attribute VB_Name = "test3"
Option Explicit
Private Att(1 To 10)
Private aType(1 To 10)
Private Const xx = "测试成功"
Private Const xy = "测试失败"
Private Const iset = "Part_info"
Private Const eklname = "sumVol"
Private Const ekldesc = "sum of vol of bodylist"
Private Const eklstr = "let lst(list) set lst=Part_info\ibodys  let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol=V"  '使用Const关键字定义常量
Sub CATmain()
    Dim oPrd, colls
    Set oPrd = CATIA.ActiveDocument.Product
    Dim refPrd: Set refPrd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refPrd.Parent.Part
  '============创建usrp参数=================

    Dim NT As Variant
    NT = Array( _
        Array("Mass", "Mass"), _
        Array("Material", "String"), _
        Array("Thickness", "Length"), _
        Array("Density", "Density") _
    )
    Dim usrp(0 To 3)
    Set colls = refPrd.UserRefProperties
    Dim i
    For i = 0 To 3
        Set usrp(i) = New Class_para
        usrp(i).SetNT NT(i)(0), NT(i)(1)
        paraDef usrp(i), colls
    Next
'============创建参数集合=================
        Set infoset = New Class_para
        infoset.SetNT "Part_info", "ParameterSet"
        Set colls = oPrt.Parameters.RootParameterSet.ParameterSets
        paraDef infoset, colls
        Set colls = infoset.obj.DirectParameters
'============创建part_info内参数=================
    Dim iBodys
    Set iBodys = New Class_para
    iBodys.SetNT "iBodys", "list"
    paraDef iBodys, colls
    If paraGetSelf(oPrt.mainbody, iBodys.obj.valuelist)(0) Is Nothing Then
            iBodys.obj.valuelist.Add oPrt.mainbody
        End If
    
    Dim infoPara(0 To 2)
     PNT = Array( _
        Array("sumVol", "Volume"), _
        Array("Thickness", "Length"), _
        Array("Density", "Density") _
    )
    For i = 0 To 2
        Set infoPara(i) = New Class_para
        infoPara(i).SetNT PNT(i)(0), PNT(i)(1)
        paraDef infoPara(i), colls
    Next
    
'============创建Relation参数=================

Set colls = oPrt.relations
Dim oRule
Set oRule = New Class_para

'        '---创建EKL
'oRule.SetNT "sum_all_Vol", "Program"
'oRule.str = eklstr
'oRule.sesc = "汇总体积"
'paraDef oRule, colls
        '---创建link_mass
oRule.SetNT "link_mass", "Formula"
oRule.Desc = "汇总体积"
oRule.Str = "Part_info\sumVol *Part_info\Density"

Set oRule.Target = refPrd.UserRefProperties.item("Mass")
paraDef oRule, colls

        '---创建link_thickness
oRule.Reset
oRule.SetNT "link_thickness", "Formula"
oRule.Desc = "链接厚度"
oRule.Str = "Part_info\Thickness"

Set oRule.Target = refPrd.UserRefProperties.item("Thickness")
paraDef oRule, colls

    
'============创建发布=================
        




'       MsgBox "不需要创建" & thispara.obj.Name & "请校验其value"
        ' select case thispara.iType
        '     Case "ParameterSet", "list"
        '         Debug.Print "不需要校验"
        '     case "Program"
        '         If thispara.obj.Value <> thispara.str Then
        '             Debug.Print "校验失败"
        '             thispara.obj.Value = thispara.str
        '             Debug.Print "校验成功"
        '     Case "Mass", "Density", "Length", "Volume" '所有dimension参数
        '         If thispara.obj.Value <> thispara.str Then
        '             Debug.Print "校验失败"
        '             thispara.obj.Value = thispara.str
        '             Debug.Print "校验成功"
        '         End If
















End Sub
Function paraDef(thispara, colls)
    If Not paraGetSelf(thispara, colls)(0) Is Nothing Then GoTo continue
        Debug.Print "需要创建" & thispara.Name
        paraCreatobj thispara, colls   '创建参数和公式时已经是默认值
        Debug.Print "已经创建" & thispara.Name
continue:
       Set thispara.obj = paraGetSelf(thispara, colls)(0) '已有参数和公式，校验默认值
       Debug.Print "不需要创建" & thispara.obj.Name

End Function
Function paraGetSelf(thispara, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.item(thispara.Name)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).Value
            paraGetSelf = arr
        Else
            paraGetSelf = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function
Function paraCreatobj(thispara, colls)
        Select Case thispara.iType
            Case "ParameterSet"
                Set thispara.obj = colls.CreateSet(thispara.Name)
            Case "list"
                Set thispara.obj = colls.CreateList(thispara.Name)
            Case "Mass", "Density", "Length", "Volume" '所有dimension参数
                Set thispara.obj = colls.CreateDimension(thispara.Name, thispara.iType, thispara.Str)
            Case "String"
                Set thispara.obj = colls.CreateString(thispara.Name, thispara.Str)
            Case "Formula"
                Set thispara.obj = colls.CreateFormula(thispara.Name, thispara.Desc, thispara.Target, thispara.Str)
            Case "Program"
                Set thispara.obj = colls.CreateProgram(thispara.Name, thispara.Desc, thispara.Str)
    End Select
End Function






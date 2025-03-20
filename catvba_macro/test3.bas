Attribute VB_Name = "test3"
Private att(1 To 10)
Private aType(1 To 10)
Private Const xx = "测试成功"
Private Const xy = "测试失败"
Private Const iset = "Part_info"
Private Const eklname = "sumVol"
Private Const ekldesc = "sum of vol of bodylist"
Private Const eklstr = "let lst(list) set lst=Part_info\ibodys  let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol=V"  '使用Const关键字定义常量
Sub CATMain()
    Set oPrd = CATIA.ActiveDocument.Product
    Dim refprd: Set refprd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refprd.Parent.Part
  '============创建usrp参数=================

    Dim NT As Variant
    NT = Array( _
        Array("Mass", "Mass"), _
        Array("Material", "String"), _
        Array("Thickness", "Length"), _
        Array("Density", "Density") _
    )

    Dim usrp(0 To 3)
    Set colls = refprd.UserRefProperties
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
    
    Dim piset(0 To 2)
     PNT = Array( _
        Array("sumVol", "Volume"), _
        Array("Thickness", "Length"), _
        Array("Density", "Density") _
    )
    For i = 0 To 2
        Set piset(i) = New Class_para
        piset(i).SetNT PNT(i)(0), PNT(i)(1)
        paraDef piset(i), colls
    Next
    
'============创建Relation参数=================
Dim oRule
Set oRule = New Class_para
oRule.SetNT "sum_all_Vol", "Program"
oRule.str = eklstr
Set colls = oPrt.relations
paraDef oRule, colls

End Sub
 Function paraDef(thispara, colls)
    If Not paraGetSelf(thispara, colls)(0) Is Nothing Then GoTo continue
        Debug.Print "需要创建" & thispara.Name
        paraCreatobj thispara, colls
        MsgBox "已经创建" & thispara.obj.Name
continue:
       Set thispara.obj = paraGetSelf(thispara, colls)(0)
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
        Select Case para.iType
            Case "ParameterSet"
                Set thispara.obj = colls.CreateSet(thispara.Name)
            Case "list"
                Set thispara.obj = colls.CreateList(thispara.Name)
            Case "Mass", "Density", "Length", "Volume" '所有dimension参数
                Set thispara.obj = colls.CreateDimension(thispara.Name, thispara.iType, thispara.str)
            Case "String"
                Set thispara.obj = colls.createstring(para.Name, para.str)
            Case "Relation"
                Set thispara.obj = colls.CreateRelation(thispara.Name, thispara.desc, thispara.Target, thispara.str)
            Case "Program"
                Set thispara.obj = colls.CreateProgram(thispara.Name, thispara.desc, thispara.str)
    End Select
End Function




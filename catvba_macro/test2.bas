Attribute VB_Name = "test2"
Private att(1 To 10)
Private aType(1 To 10)
Private Const iset = "Part_info"
Private Const eklname = "sumVol"
Private Const ekldesc = "sum of vol of bodylist"
Private Const eklstr = "let lst(list) set lst=Part_info\ibodys  let V (Volume) V=0 let i(integer) i=1 for i while i<=lst.Size() {V=V+smartVolume(lst.GetItem(i)) i=i+1} Part_info\sumVol=V"  'ʹ��Const�ؼ��ֶ��峣��


Sub test()

Dim oPrd
Set oPrd = CATIA.ActiveDocument.Product
Dim refprd: Set refprd = oPrd.ReferenceProduct
    Dim oPrt: Set oPrt = refprd.Parent.Part
'============������������=================
    Dim colls
    On Error Resume Next
        Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.item(iset).DirectParameters
        If colls Is Nothing Then
            Err.Clear
            Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.CreateSet(iset)
            Set colls = oPrt.Parameters.RootParameterSet.ParameterSets.item(iset).DirectParameters
        End If
    On Error GoTo 0
'����list����mainbody����list
    If getAtt("ibodys", colls)(0) Is Nothing Then
        Dim lst
        Set lst = colls.CreateList("ibodys")
    Else
        Set lst = getAtt("ibodys", colls)(0)
    End If
        If getAtt(oPrt.mainbody.Name, lst.valuelist)(0) Is Nothing Then
            lst.valuelist.Add oPrt.mainbody
        End If
        Set lst = Nothing
'============���������ͷ���=================
    Dim attObj
    For i = 1 To 4
        If i <> 2 Then
            If getAtt(att(i), colls)(0) Is Nothing Then
                Set attObj = colls.CreateDimension(att(i), aType(i), 0#)
            Else
                Set attObj(i) = getAtt(att(i), colls)(0)
            End If
        End If
    Next
 '=================���з���======================
    Dim Pubs
    Set Pubs = refprd.Publications
        For i = 1 To 4
            If getAtt(att(i), Pubs)(0) Is Nothing Then
                Dim oref, oPub
                Select Case i
                    Case 2
                        Set attObj(i) = refprd.UserRefProperties.item(att(i))
                End Select
            Set oref = refprd.CreateReferenceFromName(attObj(i).Name)
                Set oPub = Pubs.Add(att(i)) ' ��ӷ���
                Pubs.SetDirect att(i), oref ' ���÷���Ԫ��
            End If
        Next
        
       
        
      Set colls = oPrt.relations
      If getAtt(oFlname, colls)(0) Is Nothing Then
            Set eklobj = colls.CreateProgram(eklname, ekldesc, eklstr)
        Else
            If getAtt("ekl", colls)(1) <> eklstr Then
                getAtt("ekl", colls)(0).modify eklstr
            End If
        End If
 
        
    '================����ekl====================
    Set colls = oPrt.relations
    If getAtt("eklname", colls)(0) Is Nothing Then
        Dim eklobj, eklname, ekldesc
        eklname = "sumVol"
        ekldesc = "sum of vol of bodylist"
        Set eklobj = colls.CreateProgram(eklname, ekldesc, eklstr)
    Else
        If getAtt("ekl", colls)(1) <> eklstr Then
            getAtt("ekl", colls)(0).modify eklstr
        End If
    End If
   '================������ϵ====================
' Sub qcFormula(oPrd, item)   ' item="thickness"
    '===������ϵ===
    Dim refprd, oPrt, colls
    Dim objName, objtarget, objstr, obj
    Set refprd = oPrd.ReferenceProduct
    Set oPrt = refprd.Parent.Part
    Set colls = oPrt.relations
    Select Case item
    Case "thickness"
        objstr = "Part_info\thickness"
        objName = "Calthickness"
    Case "mass"
        objstr = "Part_info\density*Part_info\sumVol"
        objName = "Calmass"
    End If
    objtarget = refprd.UserRefProperties.item(item)
    If getAtt(objName, colls)(0) Is Nothing Then
       Set obj = colls.CreateRelation(objName, "", objtarget, objstr)
    Else
        Set obj = getAtt(objName, colls)(0)
            If obj.Value <> objstr Then
                obj.modify objstr
            End If
    End If
 End Sub
 
 
 Function getself(item, collection)
    Dim arr(1) ' ��ȷ��������
    On Error Resume Next
        Set arr(0) = collection.item(item.Name)
        If Err.Number = 0 Then ' ����Ƿ�ɹ���ȡ����
            arr(1) = arr(0).Value
            getself = arr
        Else
            getself = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function

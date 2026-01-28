Attribute VB_Name = "OTH_ivhideshow"
'Attribute VB_Name = "m64_hide&show"
'{GP:6}
'{Ep:RevertHide}
'{Caption:��ѡ����}
'{ControlTipText:��ѡ�����ؽṹ��}
'{BackColor:}
Private osel
Private rdoc
Private rPrd
Private Const mdlname As String = "OTH_ivhideshow"
Sub RevertHide()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    Set osel = pdm.msel
    Dim oDoc, cGroups, oGroup
    Set oDoc = CATIA.ActiveDocument
    Set rPrd = oDoc.Product
    Set cGroups = rPrd.GetTechnologicalObject("Groups")
    Set oGroup = cGroups.AddFromSel    ' ��ǰѡ���Ʒ���ӵ���
    oGroup.ExtractMode = 1
    oGroup.FillSelWithInvert   '  ��ѡ
    'oGroup.FillSelWithExtract
      cGroups.Remove 1
      Set cGroups = Nothing
    Dim sel
    Set sel = oDoc.Selection
    Set VisPropertySet = sel.VisProperties
    sel.VisProperties.SetShow 1  '' ������ѡ��Ԫ������Ϊ���ɼ�
    'VisPropertySet.SetShow 0
End Sub

Sub main()
    Set rdoc = CATIA.ActiveDocument
    Set rPrd = rdoc.Product
    Set osel = rdoc.Selection
    Set rPrd = CATIA.ActiveDocument.Product

Set wantlst = getwantlst
If wantlst Is Nothing Then Exit Sub
Set showlst = getshowlst(rPrd)
Set hidelst = gethidelst(rPrd)
'====��ʼ����
hide_in_lst showlst
show_in_lst wantlst
showParent wantlst
' show_in_lst showlst
'show_in_lst hidelst
'====��ʼ�ָ���ʾ

End Sub

Public Function SelectAll(iQuery As String, ByVal iRange)
  osel.Clear
  osel.Add (iRange)
  osel.Search iQuery & ",sel"
End Function
Function getshowlst(oprd)
 Dim istr As String: istr = "Assembly Design.Product.Visibility=Visible"
 Call SelectAll(istr, oprd)
  Dim lst: Set lst = KCL.InitLst
   For i = 1 To osel.count
      lst.Add osel.item(i).LeafProduct
    Next
   Set getshowlst = lst
End Function

Function gethidelst(oprd)
 Dim istr As String: istr = "Assembly Design.Product.Visibility=Hidden"
 Call SelectAll(istr, oprd)
  Dim lst: Set lst = KCL.InitLst
   For i = 1 To osel.count
      lst.Add osel.item(i).LeafProduct
    Next
   Set gethidelst = lst
End Function
Function getwantlst()
    Set getwantlst = Nothing
If osel.Count2 > 0 Then
  Dim lst: Set lst = KCL.InitLst
   For i = 1 To osel.count
      lst.Add osel.item(i).LeafProduct
    Next
        Set getwantlst = lst
End If
End Function
Sub hide_in_lst(lst)
osel.Clear
    For Each itm In lst
            osel.Add itm
    Next
osel.VisProperties.SetShow 1: osel.Clear

End Sub

Sub show_in_lst(lst)
osel.Clear
    For Each itm In lst
            osel.Add itm
    Next
osel.VisProperties.SetShow 0: osel.Clear
End Sub

Sub showParent(lst)  ' �޸�Ϊ Sub����ʵ��ʵ���߼�

    Dim parentLst: Set parentLst = KCL.InitLst()
    Dim prd, parentPrd
    
    For Each prd In lst
        Set parentPrd = prd
        ' ѭ�����ϲ���ֱ�����ڵ�
        Do
            ' ���Ի�ȡ���ڵ�
            On Error Resume Next
            Set parentPrd = parentPrd.Parent
            If Err.Number <> 0 Then
               Err.Clear
               Exit Do
            End If
            
            ' If parentPrd Is Nothing Then Exit Do ' Reached top level
            
            ' ������ڵ��� Product ���ͣ������� ReferenceProduct �� Document��
            ' ע�⣺���ݾ������ģ�Ϳ�����Ҫ�����жϣ���ͨ�� Product �� Parent ������ Product �ṹ�ڣ������
            ' �� CATIA �У�LeafProduct �� Parent ���ܻ��� Product
            If TypeName(parentPrd) = "Product" Then
                 parentLst.Add parentPrd
            ElseIf TypeName(parentPrd) = "Products" Then
                 ' ����Ǽ��ϣ���������һ��ͨ���� Product
                 Set parentPrd = parentPrd.Parent
                 If TypeName(parentPrd) = "Product" Then parentLst.Add parentPrd
            Else
                 Exit Do ' ���ﶥ���� Product �ṹ
            End If
            On Error GoTo 0
        Loop
    Next
    
    ' ͳһ��ʾ���и��ڵ�
    If parentLst.count > 0 Then
        show_in_lst parentLst
    End If
End Sub


git config --global user.name "verysolecd"
git config --global user.email "verysolecd@hotmail.com"
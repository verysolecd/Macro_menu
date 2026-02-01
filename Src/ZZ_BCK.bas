Attribute VB_Name = "ZZ_BCK"
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Const mdlName As String = "ZZ_BCK"
Sub remove_usrP()
Set oprd = CATIA.ActiveDocument.Product
rm oprd
End Sub
Sub rm(oprd)
    On Error Resume Next
     Set refPrd = oprd.ReferenceProduct
     Set oprt = refPrd.Parent.part
    Set colls = refPrd.Publications
    colls.Remove ("Location")
    colls.Remove ("iMass")
    colls.Remove ("iDensity")
    colls.Remove ("iThickness")
    colls.Remove ("iMaterial")
     Set colls = refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets
        Set cm = colls.GetItem("cm")
        Set oSel = CATIA.ActiveDocument.Selection
        oSel.Clear: oSel.Add cm: oSel.Delete
     Set colls = refPrd.Parent.part.relations
     colls.Remove ("CalM")
     colls.Remove ("CMAS")
     colls.Remove ("CTK")
     Set colls = refPrd.UserRefProperties
     colls.Remove ("iMass")
     colls.Remove ("iMaterial")
     colls.Remove ("iThickness")
    If oprd.Products.count > 0 Then
        For i = 1 To oprd.Products.count
          rm (oprd.Products.item(i))
        Next
    End If
On Error GoTo 0
End Sub


''==遍历递归=============================
Sub recurAyo(ayo)
    Dim colls: Set itm = ayo.Products
    For Each itm In colls
        Call recurFunc(itm)
    Next

    If ayo.Products.count > 0 Then
            For Each ctm In ayo.Products
                Call recurAyo(ctm)
             Next
    End If
End Sub

''==图纸页面=============================

Private Const mdlName As String = "A0_pages"
Sub main()
CATIA.RefreshDisplay = False
    Set shts = CATIA.ActiveDocument.sheets
      Set osht = Nothing
    Set lst = InitDic()
j = 1
       For i = 1 To shts.count
           Set osht = shts.item(i)
               If osht.IsDetail = False Then

                 lst.Add j, osht
        j = j + 1
               End If
       Next
    Set osht = Nothing
    For i = 1 To lst.count
       Set osht = lst(i)
       If osht.IsDetail = False Then
            osht.Activate
                    oo = straf1st(osht.Name, " ")
        If i > 9 Then
            osht.Name = "SH" & i & oo
        Else
             osht.Name = "SH0" & i & oo
        End If
            Set oView = osht.Views.item("Background View")
            Set ots = oView.Texts
            Set oDict = InitDic()
            For Each itm In ots
               Set oDict(itm.Name) = itm
            Next
           Set Pg1 = oDict("gongxxzhang")
            Pg1.text = "共" & shts.count - 1 & "页"
            Set Pg2 = oDict("dixxzhang")
            Pg2.text = "第" & i & "页"
            oView.SaveEdition
        End If
    Next
     CATIA.RefreshDisplay = True
     Set oView = osht.Views.item(1)
      osht.Activate
End Sub
Function straf1st(istr, iext)
Dim idx
idx = InStr(istr, iext)
If idx > 0 Then
        straf1st = Mid(istr, idx)
    Else
        straf1st = istr
    End If
End Function

Function InitDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = compareMode
    Set InitDic = dic
End Function







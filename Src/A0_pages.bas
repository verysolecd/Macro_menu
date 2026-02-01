Attribute VB_Name = "A0_pages"
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


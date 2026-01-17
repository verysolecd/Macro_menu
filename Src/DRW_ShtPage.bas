Attribute VB_Name = "DRW_ShtPage"
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
'            oView.Activate
            Set ots = oView.Texts
            Set oDict = InitDic()
            For Each itm In ots
               Set oDict(itm.Name) = itm
            Next
            
            Set Pg1 = oDict("gongxxzhang")
            Pg1.Text = "¹²" & shts.count - 1 & "Ò³"
            Set Pg2 = oDict("dixxzhang")
            Pg2.Text = "µÚ" & i & "Ò³"
            oView.SaveEdition
        End If
    Next
     CATIA.RefreshDisplay = True
     lst(lst.count).Activate
     Set oView = osht.Views.item(1)
End Sub
Function straf1st(iStr, iext)
Dim idx
idx = InStr(iStr, iext)
If idx > 0 Then
        straf1st = Mid(iStr, idx)
    Else
        straf1st = iStr
    End If
End Function
Function InitDic()
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.compareMode = compareMode
    Set InitDic = Dic
End Function

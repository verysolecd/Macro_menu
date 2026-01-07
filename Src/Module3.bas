Attribute VB_Name = "Module3"
Sub main()

CATIA.RefreshDisplay = False
    Set Shts = CATIA.ActiveDocument.Sheets
      Set oSHT = Nothing
    Set lst = InitDic()
j = 1
       For i = 1 To Shts.count
           Set oSHT = Shts.item(i)
               If oSHT.IsDetail = False Then

                 lst.Add j, oSHT
        j = j + 1
               End If
       Next
    Set oSHT = Nothing
    For i = 1 To lst.count
       Set oSHT = lst(i)
       If oSHT.IsDetail = False Then
            oSHT.Activate
                    oo = straf1st(oSHT.Name, " ")
        If i > 9 Then
            oSHT.Name = "SH" & i & oo

                        Else
                        oSHT.Name = "SH0" & i & oo
            End If
            Set oView = oSHT.Views.item("Background View")
'            oView.Activate
            Set ots = oView.Texts
            Set oDict = InitDic()
            For Each itm In ots
               Set oDict(itm.Name) = itm
            Next
           Set Pg1 = oDict("gongxxzhang")
            Pg1.Text = "¹²" & Shts.count - 1 & "Ò³"
            Set Pg2 = oDict("dixxzhang")
            Pg2.Text = "µÚ" & i & "Ò³"
            oView.SaveEdition
        End If
    Next
     CATIA.RefreshDisplay = True
     Set oView = oSHT.Views.item(1)
      oSHT.Activate
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


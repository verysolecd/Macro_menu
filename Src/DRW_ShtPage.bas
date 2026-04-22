Attribute VB_Name = "DRW_ShtPage"
Private Const mdlname As String = "DRW_ShtPage"
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
                    oo = StrAF(osht.Name, " ")
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
            Pg1.text = "¹²" & shts.count - 1 & "Ò³"
            Set Pg2 = oDict("dixxzhang")
            Pg2.text = "µÚ" & i & "Ò³"
            oView.SaveEdition
        End If
    Next
     CATIA.RefreshDisplay = True
     lst(lst.count).Activate
     Set oView = osht.Views.item(1)
End Sub
Function StrAF(istr, iext)
Dim idx
idx = InStr(istr, iext)
If idx > 0 Then
        StrAF = Mid(istr, idx)
    Else
        StrAF = istr
    End If
End Function
Function InitDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = compareMode
    Set InitDic = dic
End Function

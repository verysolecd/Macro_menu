Attribute VB_Name = "TEST"
Sub idn()

i = 1
For Each sht In Drawing.Sheets
    If sht.IsDetail = False Then
    sht.Parameters.item("dixxzhang").value = i
i = i + 1
End If
Next

End Sub


Attribute VB_Name = "test"

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Private mSW& ' 秒表开始时间

'Dim oDoc
'Dim CATIA, xlApp
'Public bomdata
'Public Att(1 To 4)
'Public aType(1 To 4)
Option Base 1

Sub test()
    
    Dim APP
    Set APP = GetObject(, "Excel.Application")
    Dim row_num As Long ' 声明 row_num 变量
    Dim cell ' 声明 cell 变量
    Set ws = APP.ActiveSheet
     With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
    End With
    
    For row_num = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).Value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' 设置对齐方式和缩进级别
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).row) ' 使用 xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' 使用 xlApp.xlLeft
            If IsNumeric(cell.Value) Then
                cell.IndentLevel = cell.Value ' 修改缩进级别与单元格值相等
            End If
        End If
    Next
End Sub




    



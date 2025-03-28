Attribute VB_Name = "test"

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Private mSW& ' ���ʼʱ��

'Dim oDoc
'Dim CATIA, xlApp
'Public bomdata
'Public Att(1 To 4)
'Public aType(1 To 4)
Option Base 1

Sub test()
    
    Dim APP
    Set APP = GetObject(, "Excel.Application")
    Dim row_num As Long ' ���� row_num ����
    Dim cell ' ���� cell ����
    Set ws = APP.ActiveSheet
     With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
    End With
    
    For row_num = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row ' ʹ�� xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).Value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' ���ö��뷽ʽ����������
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).row) ' ʹ�� xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' ʹ�� xlApp.xlLeft
            If IsNumeric(cell.Value) Then
                cell.IndentLevel = cell.Value ' �޸����������뵥Ԫ��ֵ���
            End If
        End If
    Next
End Sub




    



Attribute VB_Name = "Table_2_Excel_2D"
Sub CATMain()
IntCATIA
Table2Excel oActDoc
End Sub

Sub Table2Excel(oDocDrw)
If TypeName(oDocDrw) <> "DrawingDocument" Then
MsgBox "此命令只能在工程制图模块下运行!"
Exit Sub
End If

Dim drwtable As DrawingTable
Dim objWorkbook
Dim objSheet
Dim objCell
Dim InputObjectType(0)
Dim R, c
Dim IsExcelRunning As Boolean

Dim DrwSheets
Set DrwSheets = oDocDrw.Sheets


    Dim Selection
    Dim WindowLocation(1)
    Dim Status, XCenter, YCenter 'InputObjectType(0)
    Dim ObjectSelected

    Status = "MouseMove"
    InputObjectType(0) = "DrawingTable"
    ObjectSelected = False
        
    Set Selection = oDocDrw.Selection
    Selection.Clear
    
    Do While ObjectSelected = False
        Status = Selection.IndicateOrSelectElement2D("Please select drawing table to export", InputObjectType, True, True, True, ObjectSelected, WindowLocation)

    If (Status = "Cancel" Or Status = "Undo" Or Status = "Redo") Then
        MsgBox "你没有选择表格，程序退出!"
        Exit Sub
    End If
       
    Loop
    
    If ObjectSelected Then
       Set drwtable = Selection.Item(1).Value
    End If
    
Dim totalrows, totalcolumns
totalrows = drwtable.NumberOfRows
totalcolumns = drwtable.NumberOfColumns
Dim table()
ReDim table(totalrows, totalcolumns)
For R = 1 To totalrows
For c = 1 To totalcolumns
table(R, c) = drwtable.GetCellString(R, c)
Next
Next

IntExcel

objExcel.Visible = True

Dim myname
myname = Replace(oDocDrw.FullName & "_" & drwtable.Name, ".", "_") & ".xlsx"
Set objWorkbook = objExcel.workbooks.Add()
objWorkbook.Activate
Set objSheet = objWorkbook.Worksheets.Item(1)
'--Populate spreadsheet with values from array.
For R = 1 To totalrows
For c = 1 To totalcolumns
objSheet.cells(R, c) = table(R, c)
Next
Next
On Error Resume Next
objWorkbook.SaveAs myname
If Err.Number <> 0 Then
  MsgBox "没有权限保存文件" & vbNewLine & vbNewLine & myname & vbNewLine & vbNewLine & "请关闭已打开的文件或使用其他文件名手动另存", vbInformation, "Save excel failure"
End If
On Error GoTo 0
'objWorkbook.Close
'objExcel.Quit ' Close Excel with the Quit method on the Application object
'MsgBox "Your Catia drawing table has been exported to Excel." & vbNewLine & vbNewLine & myname, vbInformation, "Table Export Complete"
End Sub


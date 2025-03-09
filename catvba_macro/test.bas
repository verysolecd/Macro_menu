Attribute VB_Name = "test"

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Private mSW& ' 秒表开始时间
Option Explicit '初始化对象
Dim oDoc
Dim CATIA, xlApp
Public counter
Public bomdata
Public att(1 To 4)
Public aType(1 To 4)

Sub test()
 Dim oDoc, iPrd, rootPrd, oPrd, children, oPrt, refprd
     Dim xlsht, startrow, startcol, currRow, LV, rng
     Dim propertyArry()
     Dim i
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
           Set oDoc = CATIA.ActiveDocument
    Set rootPrd = CATIA.ActiveDocument.Product
         If Err.Number <> 0 Then
            MsgBox "请打开CATIA并打开你的产品，再运行本程序": Err.Clear
            Exit Sub
         End If
    On Error GoTo 0
    
  
   CATIA.ActiveWindow.WindowState = 0
   CATIA.Visible = True
       Set xlApp = GetObject(, "Excel.Application")
    Set xlsht = xlApp.ActiveSheet
   iniarr
   
   Dim arry
   arry = recurPrd(rootPrd, 0)
   
    Dim fn
    fn = counter
    With xlsht
    xlsht.Range(.Cells(2, 1), .Cells(fn + 1, 11)).Value = arry
  End With
End Sub
 


    



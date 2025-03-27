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



    Dim bom_cols
    bom_cols = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    
    ' 输出数组的索引和对应元素
    For i = LBound(bom_cols) To UBound(bom_cols)
        Debug.Print "Index: " & i & ", Value: " & bom_cols(i)
    Next i
End Sub





    



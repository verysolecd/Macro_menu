Attribute VB_Name = "test"

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Private mSW& ' 秒表开始时间

Dim oDoc
Dim CATIA, xlApp

Public bomdata
Public Att(1 To 4)
Public aType(1 To 4)


Sub test()





End Sub

    



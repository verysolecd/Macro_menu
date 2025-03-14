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

    
    Dim att(1 To 4)
    att(1) = "iMaterial"
    att(2) = "iDensity"
    att(3) = "iMass"
    att(4) = "iThickness"
    
Dim xlm, pdm
Set xlm = New Class_XLM
Set pdm = New class_PDM

xlm.inject_data 1, att
  
End Sub
 


    



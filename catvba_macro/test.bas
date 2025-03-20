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

Dim me As New Class_para

    Set pdm = New class_PDM
     

'    att(1) = "iMass"
'    att(2) = "iMaterial"
'    att(3) = "iThickness"
'    att(4) = "iDensity"
'    aType(1) = "Mass"
'    aType(2) = "String"
'    aType(3) = "Length"
'    aType(4) = "Density"

 Dim i
 
 i = 1
    att(i).Name = "Mass"
    att(i).iType = "Mass"
    att(i).Value = 0#
   Set att(i).Target = rootPrd
    
MsgBox att(i).Name & att(i).iType & att(i).Value & att(i).Target.Name


End Sub


    



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
 


Function recurPrd(oPrd, LV)

     If counter = 0 Then
          ReDim bomdata(1 To 1000, 1 To 11) ' 扩展为11列：2列原有数据 + 9列产品属性
     '     If IsEmpty(bomdata) Then
    End If
    counter = counter + 1
    bomdata(counter, 1) = counter
    bomdata(counter, 2) = LV
    
    Dim prdInfo, j
     prdInfo = infoPrd(oPrd)
     For j = 1 To 9
         bomdata(counter, j + 2) = prdInfo(j)
     Next j
    
    Dim children As Products
    Set children = oPrd.Products
    If children.Count > 0 Then
        Dim i As Integer
        For i = 1 To children.Count
            recurPrd children.Item(i), LV + 1
        Next
    End If
    recurPrd = bomdata

End Function



Function count_me(oPrd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
     If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
               Dim oParent: Set oParent = oPrd.Parent.Parent
         
              Set oDict = CreateObject("Scripting.Dictionary")
              For i = 1 To oParent.Products.Count
                     pn = oParent.Products.Item(i).PartNumber
                     If oDict.Exists(pn) = True Then
                         oDict(pn) = oDict(pn) + 1
                     Else
                         oDict(pn) = 1
                     End If
                 Next
        QTy = oDict(oPrd.PartNumber)       '获取oprd数量
     End If
     If Error.Number <> 0 Then
          QTy = 1
     End If
    count_me = QTy
End Function



Function infoPrd(oPrd)
        Dim arr(1 To 9)
            With oPrd.ReferenceProduct
                arr(1) = .PartNumber
                arr(2) = .Nomenclature
                arr(3) = .Definition
                arr(4) = oPrd.Name
            End With
        Dim usrp
           Set usrp = oPrd.ReferenceProduct.UserRefProperties
                arr(5) = getAtt("iMass", usrp)(1)
                arr(6) = getAtt("iMaterial", usrp)(1)
                arr(7) = getAtt("iThickness", usrp)(1)
        On Error Resume Next
           Set usrp = oPrd.ReferenceProduct.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item("cm").DirectParameters
                arr(8) = getAtt("iDensity", usrp)(1)
            If Error.Number <> 0 Then
                arr(8) = "__"
            End If
        On Error GoTo 0
                arr(9) = count_me(oPrd)
        infoPrd = arr()
    End Function
Sub iniarr()

    att(1) = "iMaterial"
    att(2) = "iDensity"
    att(3) = "iMass"
    att(4) = "iThickness"
    aType(1) = "String"
    aType(2) = "Density"
    aType(3) = "Mass"
    aType(4) = "Length"
End Sub

    
Function getAtt(itemName, collection)
    Dim arr(1) ' 正确声明数组
    On Error Resume Next
        Set arr(0) = collection.Item(itemName)
        If Err.Number = 0 Then ' 检查是否成功获取对象
            arr(1) = arr(0).Value
            getAtt = arr
        Else
            getAtt = Array(Nothing, "__")
        End If
    On Error GoTo 0
End Function


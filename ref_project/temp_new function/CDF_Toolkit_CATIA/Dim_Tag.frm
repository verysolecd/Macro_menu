VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dim_Tag 
   Caption         =   "Dimension_Tag"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   OleObjectBlob   =   "Dim_Tag.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dim_Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAuto_Click()

Set oSel = oActDoc.Selection
        oSel.Clear
        Dim ExistingNums As New VBA.Collection
        Set ExistingNums = colSearchName("MyDim_")
        

        If ExistingNums.Count <> 0 Then
            If MsgBox("将删除已有的" & ExistingNums.Count & " 个尺寸编号,再重新生成！" & vbCrLf & "点击确定继续，点击取消退出", 1) = vbOK Then
                      Dim i
                         For Each i In ExistingNums
                         oSel.Add i
                         Next
                      On Error Resume Next
                      oSel.Delete
                        If Err.Number <> 0 Then
                              MsgBox "删除失败! 请检查是否有视图被锁定,如有则请先解除视图锁定"
                              Exit Sub
                        End If
                      On Error GoTo 0
                      txtNextno.Value = 1
            Else
               Exit Sub
            End If
         End If
  
AutoNum 1
RefreshInfo
'txtNextno.Value = sFinNum + 1
End Sub

Private Sub cmdDelAll_Click()
On Error Resume Next
Err.Clear
Set oSel = oActDoc.Selection
        oSel.Clear
        Dim ExistingNums As New VBA.Collection
        Set ExistingNums = colSearchName("MyDim_")

        If ExistingNums.Count <> 0 Then
            If MsgBox(ExistingNums.Count & " 个尺寸编号将被删除", 1) = vbOK Then
                      Dim i
                         For Each i In ExistingNums
                         oSel.Add i
                         Next
                      oSel.Delete
                      txtNextno.Value = 1
            Else
               Exit Sub
            End If
        Else
            MsgBox "没有此程序生成的尺寸编号"
            GoTo dNote
            Exit Sub
        End If
        
'删除尺寸编号注释
dNote:
oSel.Clear
        Dim ExistingNums1 As New VBA.Collection
        Set ExistingNums1 = colSearchName("MyNumNote")
        If ExistingNums1.Count <> 0 Then
                      Dim i1
                         For Each i1 In ExistingNums1
                         oSel.Add i1
                         Next
                      oSel.Delete
        End If
If Err.Number <> 0 Then
MsgBox "删除失败! 请检查是否有视图被锁定,如有则请先解除视图锁定"
End If
On Error GoTo 0
RefreshInfo
End Sub

Private Sub cmdExport_Click()
Dim ActDocument As DrawingDocument
Dim DrawingSheets2 As DrawingSheets
Dim drawingSheet2 As DrawingSheet
Dim drawingViews2 As DrawingViews
Dim drawingView2 As DrawingView
Dim DrwText2

Set ActDocument = CATIA.ActiveDocument


If (TypeName(ActDocument) <> "DrawingDocument") Then
    MsgBox "This macro runs in CATDrawing env.", vbCritical, "Information"
    Exit Sub
End If


Set DrawingSheets2 = ActDocument.Sheets

Dim int_i, int_j, int_k, int_x
Dim ArrMyDimNums()

ReDim ArrMyDimNums(0)

    For int_i = 1 To DrawingSheets2.Count

        Set drawingSheet2 = DrawingSheets2.Item(int_i)

                For int_j = 1 To drawingSheet2.Views.Count

                     Set drawingView2 = drawingSheet2.Views.Item(int_j)

                        For int_k = 1 To drawingView2.Texts.Count

                            Set DrwText2 = drawingView2.Texts.Item(int_k)

'                                If Left(DrwText2.Name, Len(PreFix)) = PreFix Then
                                If Left(DrwText2.Name, 6) = "MyDim_" Then

                                    int_x = UBound(ArrMyDimNums())
                              
                                    ReDim Preserve ArrMyDimNums(int_x + 1)

                                    Set ArrMyDimNums(int_x + 1) = DrwText2

                                End If

                        Next

                Next

    Next

MsgBox "The Total Dimension Identification Number is " & UBound(ArrMyDimNums)

Dim table1, table2


table1 = ActDocument.Name & "尺寸编号报告" & vbCrLf & _
                "臭豆腐工具箱CATIA版于" & Now() & "自动生成" & vbCrLf & vbCrLf & _
                "尺寸编号No.," & "尺寸Dimension,规格Specification,尺寸极限DimLimit," & "下公差LowTol," & "上公差UpTol," & "备注"

table2 = ""
'给数组ArrMyDinNums排序
Dim temp1
int_x = UBound(ArrMyDimNums)
For int_i = 1 To int_x
    For int_j = 1 To int_x - int_i
        If CInt(ArrMyDimNums(int_j).Text) > CInt(ArrMyDimNums(int_j + 1).Text) Then
           Set temp1 = ArrMyDimNums(int_j)
           Set ArrMyDimNums(int_j) = ArrMyDimNums(int_j + 1)
           Set ArrMyDimNums(int_j + 1) = temp1
        End If
    Next int_j
Next int_i


'给数组ArrMyDinNums排序

Dim oMyDim
For int_x = 1 To UBound(ArrMyDimNums)
On Error Resume Next
   Set oMyDim = ArrMyDimNums(int_x).AssociativeElement
   'MsgBox ArrMyDimNums(int_x).Text & "# TypeName is  " & TypeName(oMyDim) & vbCrLf & "DimType is " & oMyDim.DimType & vbCrLf & vbCrLf & "DimValue Status is  " & oMyDim.DimStatus
   
   
   If Err.Number <> 0 Then
       table2 = ArrMyDimNums(int_x).Text & "," & "没有关联对象,,," & "," & "," & ","
   Else
       Dim valuesA
       If TypeName(oMyDim) = "DrawingDimension" Then
       
          valuesA = GetQCPValue(oMyDim)
          table2 = ArrMyDimNums(int_x).Text & "," & valuesA(0) & "," & valuesA(1) & "," & valuesA(2) & "," & valuesA(3) & "," & valuesA(4) & "," & valuesA(5)
       Else
          table2 = ArrMyDimNums(int_x).Text & "," & "关联对象不是线性尺寸,,," & "," & "," & ","
       
       End If
    End If
    
    table1 = table1 & vbCrLf & table2
Next
CreateTxt ActDocument.Name, table1
'MsgBox table1
End Sub

Private Sub cmdManual_Click()
ManualNo txtNextno.Value
txtNextno.Value = txtNextno.Value + 1
End Sub

Private Sub cmdNote_Click()
AddNumNote
RefreshInfo
End Sub

Sub cmdRefresh_Click()
RefreshInfo
End Sub

Private Sub cmdSelAll_Click()
Set oSel = oActDoc.Selection
        oSel.Clear
        Dim ExistingNums As New VBA.Collection
        Set ExistingNums = colSearchName("MyDim_")
        

        If ExistingNums.Count <> 0 Then
                    Dim i
                         For Each i In ExistingNums
                         oSel.Add i
                         Next
                    MsgBox oSel.Count & " 个尺寸编号已被选择"
         Else
               MsgBox "没有此程序生成的尺寸编号"
         End If

End Sub
Private Sub cmdSort_Click()
SortAll
RefreshInfo
End Sub



Private Sub CurrFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Open_Current_Folder.CATMain
End Sub

Private Sub lblNextno_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
txtNextno.Value = sFinNum + 1
End Sub
Private Sub lstNumbers_Click()

IntCATIA
Set oSel = oActDoc.Selection
oSel.Clear
Dim i, j, sNumber

sNumber = vbNullString
For i = 0 To lstNumbers.ListCount - 1
If lstNumbers.Selected(i) = True Then
sNumber = lstNumbers.List(i)
'MsgBox "选择了 " & sNumber
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
End If
Next

Dim ExistingNums As New VBA.Collection
Set ExistingNums = colSearchName("MyDim_" & "s" & sNumber)
        If ExistingNums.Count <> 0 Then
                    Dim Nums
                         For Each Nums In ExistingNums
                         oSel.Add Nums
                         Next
         End If

End Sub

Private Sub lstNumbers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
lstNumbers_Click
Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = CATIA.ActiveWindow

Dim specsViewer1 As SpecsViewer
Set specsViewer1 = specsAndGeomWindow1.ActiveViewer

specsViewer1.Reframe
End Sub

Sub CreateTxt(sTextFileName, sText)
   Dim objShell
   Set objShell = CreateObject("Shell.Application")
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim FoldPath, FoldObj
   Set FoldPath = objShell.BrowseForFolder(0, "Define a Folder:", 0, 0)
   If FoldPath Is Nothing Then
  MsgBox "You didn't choose a folder, Macros exit !" & vbNewLine & "你没有选择文件夹，程序退出！", vbExclamation, "DensityVolumeMass"
  Exit Sub
  Else
 Set FoldObj = objFSO.GetFolder(FoldPath.Self.Path)
 End If

Dim objStream, FILE_NAME
Const TristateFalse = 0
Randomize
FILE_NAME = FoldPath.Self.Path & "\" & sTextFileName & "_" & Int(100 * Rnd) & ".csv"
Set objStream = objFSO.CreateTextFile(FILE_NAME, True, TristateFalse)
objStream.Write (sText)
objStream.Close
MsgBox "数据保存于 " & FILE_NAME
End Sub
Function GetQCPValue(oDrawingDimension)
On Error Resume Next
Dim ArrT0, ArrT1, ArrT2, ArrT3, ArrT4, remark0

ArrT0 = ""
ArrT1 = ""
ArrT2 = ""
ArrT3 = ""
ArrT4 = ""
remark0 = ""

Dim oDimValue
Dim oPrefix, oSuffix
Dim iBefore, iAfter, iUpper, iLower
Dim oTolType, oTolName, oUpTol, oLowTol, odUpTol, odLowTol, oDisplayMode
 
Set oDimValue = oDrawingDimension.GetValue
    If Err.Number <> 0 Then
       oPrefix = ""
       oSuffix = ""
       iBefore = ""
       iAfter = ""
       iUpper = ""
       iLower = ""
    Else
       oDimValue.GetPSText 1, oPrefix, oSuffix    '1: main value 2: dual value,prefix text.SufFix Text
       oDimValue.GetBaultText 1, iBefore, iAfter, iUpper, iLower
    End If
    oPrefix = Replace(oPrefix, "<DIAMETER>", "D")
    
oDrawingDimension.GetTolerances oTolType, oTolName, oUpTol, oLowTol, odUpTol, odLowTol, oDisplayMode
If Err.Number <> 0 Then
    oTolType = ""
    oTolName = ""
    oUpTol = ""
    oLowTol = ""
    odUpTol = ""
    odLowTol = ""
    oDisplayMode = ""
End If



If oDrawingDimension.ValueFrame = 5 Then
    'remark0 = remark0 & "/理论正确尺寸TED"
    ArrT2 = "理论正确尺寸TED"
End If

If oDrawingDimension.DimStatus = 2 Then
   ArrT0 = CStr(oDimValue.GetFakeDimValue(1))
   remark0 = remark0 & "/假尺寸FakeDim"
Else
    Dim Precision
    Precision = Len(oDimValue.GetFormatPrecision(1)) - InStr(oDimValue.GetFormatPrecision(1), ".")
    ArrT0 = CStr(Round(oDimValue.Value, Precision))
    
    If oDrawingDimension.DimType = 4 Then
        ArrT0 = CStr(Round(oDimValue.Value * 180 / 3.1415926535, Precision))
    End If

End If

ArrT1 = Replace(iBefore & " " & oPrefix & " " & ArrT0 & " " & oSuffix & " " & iAfter & " " & iUpper & " " & iLower, ",", "_")

If oDrawingDimension.ValueFrame <> 5 Then
        If odLowTol = 0 And odUpTol = 0 Then
            ArrT2 = ""
            ArrT3 = ""
            ArrT4 = ""
        Else
            ArrT2 = Val(ArrT0) + odLowTol & " to " & Val(ArrT0) + odUpTol
            ArrT3 = odLowTol
            ArrT4 = odUpTol
        End If
End If

If oDrawingDimension.DimType = 4 Then
    ArrT2 = "角度 " & ArrT2
End If

GetQCPValue = Array(ArrT0, ArrT1, ArrT2, ArrT3, ArrT4, remark0)

End Function

Private Sub UserForm_Initialize()
On Error Resume Next
RefreshInfo

SetHotkey 1, 115, "Add", "Dimension_Tag" '按 F4 激活指定程序，F4的Ascii码为115
SetHotkey 2, 116, "Add", "Dimension_Tag" '按 F5 激活指定程序，F5的Ascii码为116
'SetHotkey 3, 32, "Add", "Dimension_Tag" '按空格激活指定程序，空格的Ascii码为32
End Sub


Private Sub UserForm_Terminate()
SetHotkey 1, "", "Del", "Dimension_Tag" '取消 F4热键，F4的Ascii码为115
SetHotkey 2, "", "Del", "Dimension_Tag" '取消 F5热键，F4的Ascii码为116
'SetHotkey 3, "", "Del", "Dimension_Tag" '取消 空格热键，F12的Ascii码为32
End Sub

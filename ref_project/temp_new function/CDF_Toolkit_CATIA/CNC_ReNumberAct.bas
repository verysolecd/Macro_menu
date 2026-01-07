Attribute VB_Name = "CNC_ReNumberAct"
Option Explicit
    Dim ToolChangeNo As Integer
    Dim TableHeadRotationNo As Integer
    Dim CoordinateSystemNo As Integer
    Dim PPInstructionNo As Integer
    Dim ActNo As Integer

Sub CATMain()

       
IntCATIA
If TypeName(oActDoc) <> "ProcessDocument" Then
    MsgBox "此命令只能在加工模块中运行！"
    Exit Sub
Else
    Dim oMyPPR As Object
    Set oMyPPR = oActDoc.PPRDocument
End If
ToolChangeNo = 0
TableHeadRotationNo = 0
CoordinateSystemNo = 0
PPInstructionNo = 0
ActNo = 0

Dim Sel2 As Object
On Error Resume Next
Set Sel2 = MProg.Sel("ManufacturingSetup", "ManufacturingProgram")
        If Err.Number <> 0 Then
        Exit Sub
        End If

    Dim oMyPrgs As Object
    Dim iii As Integer
    Select Case TypeName(Sel2)
            Case "ManufacturingSetup"
                Set oMyPrgs = Sel2.Programs
                For iii = 1 To oMyPrgs.Count
                     Dim oMyPrg As Object
                     Set oMyPrg = oMyPrgs.GetElement(iii)
                         'MsgBox oMyPrg.Name & ",TypeName is " & TypeName(oMyPrg)
                         If oMyPrg.Active Then
                         ReNumb oMyPrg, False
                         End If
                Next 'iii
            Case "ManufacturingProgram"
                ReNumb Sel2, True
    End Select

'MsgBox "操作完成，请按Ctrl+S 保存修改", vbInformation, "行为重新排序"

End Sub

Sub ReNumb(mp As Object, ActOnly As Boolean) 'mp is program

    If TypeName(mp) <> "ManufacturingProgram" Then
        Exit Sub
    End If
On Error Resume Next
Dim Acts As Object
Dim iiii As Integer
Set Acts = mp.Activities

  For iiii = 1 To Acts.Count
   Dim Act As Object
   Set Act = Acts.GetElement(iiii)

   If Act.Active Then
'      If (Act.Type <> "ToolChange" And Act.Type <> "TableHeadRotation" And Act.Type <> "CoordinateSystem" And Act.Type <> "PPInstruction") Then
        Select Case Act.Type
            Case "ToolChange"
                  If ActOnly = False Then
                  ToolChangeNo = ToolChangeNo + 1
                  Act.Name = TrailNumb(Act.Name, True, ToolChangeNo, 2)
                  End If
            Case "TableHeadRotation"
                  If ActOnly = False Then
                  TableHeadRotationNo = TableHeadRotationNo + 1
                  Act.Name = TrailNumb(Act.Name, True, TableHeadRotationNo, 2)
                  End If
            Case "CoordinateSystem"
                  If ActOnly = False Then
                  CoordinateSystemNo = CoordinateSystemNo + 1
                  Act.Name = TrailNumb(Act.Name, True, CoordinateSystemNo, 2)
                  End If
            Case "PPInstruction"
                  If ActOnly = False Then
                  PPInstructionNo = PPInstructionNo + 1
                  Act.Name = TrailNumb(Act.Name, True, PPInstructionNo, 2)
                  End If
            Case Else
                  ActNo = ActNo + 1
                  Act.Name = TrailNumb(Act.Name, True, ActNo, 3)
        End Select
        
   End If
  Next 'iiii
End Sub
Function TrailNumb(ByVal S0 As String, TrimTrail As Boolean, Optional ByVal n1 As String, Optional ByVal MinDigits As Integer)
If TrimTrail And InStr(S0, ".") > 1 Then
TrailNumb = Mid(S0, 1, InStr(S0, ".") - 1)
Else
TrailNumb = S0

End If

If MinDigits > 1 And MinDigits < 5 Then
    Do While Len(n1) < CInt(MinDigits)
         n1 = "0" & CStr(n1)
    Loop
Else
MsgBox "输入的序号位数应该在2,3,4之间！"
End If

TrailNumb = TrailNumb & "." & n1

End Function

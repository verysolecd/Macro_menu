VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MProg 
   Caption         =   "MP_ReName_Export"
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "MProg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DicMPs As Object  'FSO directionary
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Private Sub cmdClear_Click()
DicMPs.RemoveAll
lbxMPs.Clear

cmdReName.Enabled = True
cmdReName2.Enabled = True
txtPreFix.Enabled = True
txtStartNum.Enabled = True
txtStep.Enabled = True

lbxMPs.Enabled = False

cmdClear.Enabled = False

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
Dim i As Integer
For i = 0 To lbxMPs.ListCount - 1
If lbxMPs.Selected(i) Then
   If DicMPs.Exist(lbxMPs.List(i, 0)) Then
        DicMPs.Remove (lbxMPs.List(i, 0))
   End If
End If
Next
lbxMPs.Clear
lbxMPs.List = DicMPs.keys


' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
End Sub

Private Sub cmdGen_Click()
        
        If lbxMPs.ListCount = 0 Then
        MsgBox "请先选择MP程序！"
        Exit Sub
        End If
cmdGen.Caption = "正在生成..."
Dim i, sResult, iFail
iFail = 0

For i = 0 To lbxMPs.ListCount - 1
sResult = "-成功"
On Error Resume Next
CreateNCfile DicMPs.Item(lbxMPs.List(i, 0)), cmbx_Type.Text
If Err.Number <> 0 Then
   sResult = "-失败!!"
   iFail = iFail + 1
End If
lbxMPs.List(i, 0) = lbxMPs.List(i, 0) & sResult
On Error GoTo 0
Next


        If iFail = 0 Then
          
             If MsgBox("操作了 " & lbxMPs.ListCount & " 个MP程序,全部成功！" _
                        & vbCrLf & "文件存放于 " & oCATVBA_Folder("NC_Files").Path & "\ 目录", vbOKCancel) = vbOK Then
                Shell "explorer.exe " & oCATVBA_Folder("NC_Files").Path, vbNormalFocus
             End If
            
        Else

             If MsgBox("操作了 " & lbxMPs.ListCount & " 个MP程序, " & iFail & " 个失败！" _
                        & vbCrLf & "文件存放于 " & oCATVBA_Folder("NC_Files").Path & "\ 目录", vbOKCancel) = vbOK Then
                Shell "explorer.exe " & oCATVBA_Folder("NC_Files").Path, vbNormalFocus
             End If
        End If
DicMPs.RemoveAll
cmdGen.Caption = "生成刀轨文件"
On Error Resume Next
Kill oCATVBA_Folder("NC_Files").Path & "\*.LOG"
Kill oCATVBA_Folder("NC_Files").Path & "\*_LOG"

End Sub

Private Sub cmdPrgToolList_Click()
CNC_Prog_List.CATMain
End Sub

Private Sub cmdReName_Click()

On Error Resume Next
Dim i, MPNum, MPNumStep
MPNum = CLng(txtStartNum)
MPNumStep = CInt(txtStep)

    If Err.Number <> 0 Then
    MsgBox "输入错误"
    Exit Sub
    End If
    
    If CInt(txtStep) = 0 Then
    MsgBox "步长不能为0"
    Exit Sub
    End If
On Error GoTo 0

On Error Resume Next
Dim digits As String
digits = ""
For i = 1 To Len(txtStartNum)
digits = CStr(digits & "0")
Next

Dim op As ManufacturingSetup
Dim MPs As MfgActivities
Set op = Sel("ManufacturingSetup")
If Err.Number <> 0 Then
Exit Sub
End If
Set MPs = op.Programs

For i = 1 To MPs.Count
If MPs.GetElement(i).Active Then
MPs.GetElement(i).Name = txtPreFix & Format(MPNum, digits) & txtSufFix
MPNum = MPNum + MPNumStep
End If
Next
txtPreFix.Enabled = True
txtStartNum.Enabled = True
txtStep.Enabled = True

MsgBox "操作完成！请 Ctrl + S 保存修改"

End Sub
Function Sel(objType As String, Optional ObjType2 As String, Optional ObjType3 As String)
Dim s2, InputObjectType(), Status, NotSel, obj, Picked
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear

If ObjType2 = "" Then
        ReDim InputObjectType(0)
        InputObjectType(0) = objType
        ElseIf ObjType3 = "" Then
                ReDim InputObjectType(1)
                InputObjectType(0) = objType
                InputObjectType(1) = ObjType2
                Else
                    ReDim InputObjectType(2)
                    InputObjectType(0) = objType
                    InputObjectType(1) = ObjType2
                    InputObjectType(2) = ObjType3
End If

Dim indication, n
indication = ""
For n = 0 To UBound(InputObjectType)
    indication = indication & "," & InputObjectType(n)
Next

Picked = False
NotSel = True
    Do While NotSel
        Status = s2.SelectElement2(InputObjectType, "Select the " & indication, False)
        If (Status = "Cancel") Then
            Exit Function
        ElseIf (Status = "Redo" And Not Picked) Then
               ElseIf (Status = "Undo") Then
                        Exit Function
                    ElseIf (Status <> "Redo") Then Set obj = s2.Item(1).Value
                              Picked = True
                              NotSel = False
        End If
    Loop
    s2.Clear
    Set Sel = obj
End Function

Private Sub cmdReName2_Click()
ReplaceNum.Show vbModeless
ReplaceNum.Left = ReplaceNum.Left * 2
End Sub

Private Sub cmdReNameAct_Click()
CNC_ReNumberAct.CATMain

End Sub

Sub cmdSel_Click()
cmdReName.Enabled = False
cmdReName2.Enabled = False
txtPreFix.Enabled = False
txtStartNum.Enabled = False
txtStep.Enabled = False

lbxMPs.Enabled = True


Dim Sel2 As Object
On Error Resume Next
Set Sel2 = Sel("ManufacturingSetup", "ManufacturingProgram")
        If Err.Number <> 0 Then
        Exit Sub
        End If
'On Error GoTo 0

If TypeName(Sel2) = "ManufacturingProgram" Then
           If DicMPs.exists(Sel2.Name) = False And Sel2.Active Then
              DicMPs.Add Sel2.Name, Sel2
              lbxMPs.AddItem Sel2.Name
            End If
        ElseIf Sel2.Programs.Count <> 0 Then
              Dim ii
              For ii = 1 To Sel2.Programs.Count
                         If DicMPs.exists(Sel2.Programs.GetElement(ii).Name) = False _
                          And Sel2.Programs.GetElement(ii).Active Then
                              DicMPs.Add Sel2.Programs.GetElement(ii).Name, Sel2.Programs.GetElement(ii)
                              lbxMPs.AddItem Sel2.Programs.GetElement(ii).Name
                         End If
              Next
End If
cmdClear.Enabled = True
cmdDel.Enabled = True
End Sub
Sub CreateNCfile(mp As ManufacturingProgram, sDataType As String)
Dim outputGen As ManufacturingOutputGenerator
Dim genData As ManufacturingGeneratorData

Dim OutPutFolder As Object
Dim FileName As String
Dim FullFileName As String

Set OutPutFolder = oCATVBA_Folder("NC_Files")

Select Case sDataType
       Case "APT"
            'FileName = Replace(oActDoc.Name, ".CATProcess", "-") & MP.Name & ".aptsource"
            FileName = mp.Name & ".aptsource"  '名称不含文件名
       Case "CLF"
            'FileName = Replace(oActDoc.Name, ".CATProcess", "-") & MP.Name & ".clfile"
            FileName = mp.Name & ".clfile" '名称不含文件名

       Case Else
End Select

FullFileName = OutPutFolder.Path & "\" & FileName

Set outputGen = mp
outputGen.InitFileGenerator sDataType, FullFileName, genData
outputGen.RunFileGenerator genData
genData.ResetAllModalValues
End Sub


Private Sub CurrFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Open_Current_Folder.CATMain
End Sub

Private Sub lbl_Type_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Shell "explorer.exe " & oCATVBA_Folder("NC_Files").Path, vbNormalFocus
On Error GoTo 0
End Sub

Private Sub txtStartNum_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入数字！"
    End Select
  End Sub


Private Sub txtStep_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入数字！"
    End Select
End Sub


Private Sub UserForm_Initialize()
IntCATIA

cmbx_Type.AddItem "APT"
cmbx_Type.AddItem "CLF"

lbxMPs.Enabled = False
cmdClear.Enabled = False
cmdDel.Enabled = False
Set DicMPs = CreateObject("Scripting.Dictionary") 'MP list

Dimension_Tag_2D.SetHotkey 3, 120, "Add", "MP_ReName_Export" '按F9激活指定程序，F9的Ascii码为120
ReadConf
End Sub


Private Sub ReadConf()
Dim sConfpath As String
sConfpath = oCATVBA_Folder.Path
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'MsgBox sConfpath
'MsgBox Dir(sConfpath)
If Not objFSO.FileExists(sConfpath & "\Conf1.ini") Then
Exit Sub
End If
'---读取配置文件---
'On Error Resume Next
    Dim read_OK ' As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("程序批量改名", "程序名前缀", "O10", read2, 256, sConfpath & "\Conf1.ini")
    txtPreFix.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("程序批量改名", "程序名后缀", "", read2, 256, sConfpath & "\Conf1.ini")
    txtSufFix.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("程序批量改名", "DataType", "APT", read2, 256, sConfpath & "\Conf1.ini")
    cmbx_Type.Text = read2

End Sub
Private Sub SaveConf()
On Error Resume Next
Dim sConfpath As String
Dim write1 'As Long
    '参数一： Section Name (节的名称)。
    '参数二： 节下面的项目名称。
    '参数三： 项目的内容。
    '参数四： ini配置文件的路径名称。
              
sConfpath = oCATVBA_Folder.Path
write1 = WritePrivateProfileString("程序批量改名", "程序名前缀", txtPreFix.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("程序批量改名", "程序名后缀", txtSufFix.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("程序批量改名", "DataType", cmbx_Type.Text, sConfpath & "\Conf1.ini")

End Sub

Private Sub UserForm_Terminate()
Dimension_Tag_2D.SetHotkey 3, "", "Del", "MP_ReName_Export" '取消热键
Set DicMPs = Nothing

SaveConf
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Batch_Prnt 
   Caption         =   "臭豆腐工具箱CATIA版 | 批量输出"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "Batch_Prnt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Batch_Prnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOutFolder As Object
Dim DicDwgs As Object
Dim sCap As String



Private Sub cmdOutDir_Click()
On Error Resume Next
 Dim oFileSys
   Dim FoldObj
   Set oFileSys = CreateObject("Scripting.FileSystemObject") ' CATIA.FileSystem
   Dim objShell
   Set objShell = CreateObject("Shell.Application")
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim FoldPath
   Set FoldPath = objShell.BrowseForFolder(0, "请选择一个输出目录:", 0, 0)
       'objShell.CascadeWindows
   If FoldPath Is Nothing Then
'      MsgBox "你没有选择文件夹，当前输出目录为" & vbCrLf & _
'      oOutFolder.Path, , "批量导出"
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
      Exit Sub   'WScript.Quit
  Else
  Set oOutFolder = objFSO.GetFolder(FoldPath.Self.Path)
  On Error GoTo 0
  lblOut1.Caption = oOutFolder.Path
 End If
End Sub

Private Sub cmdRunBatch_Click()
        
        If lbxFiles.ListCount = 0 Then
        MsgBox "请先选择图纸！", , "批量导出"
        Exit Sub
        End If
CATIA.DisplayFileAlerts = False
CATIA.RefreshDisplay = False

cmdRunBatch.Caption = sCap & "-->>>正在进行"
cmdRunBatch.BackColor = &H80000000
Dim i, sResult, iFail
iFail = 0




For i = 0 To lbxFiles.ListCount - 1
        sResult = ""
        lbxFiles.Selected(i) = True
    On Error Resume Next
        Dim oDwgi As Object
        Set oDwgi = CATIA.Documents.Read(DicDwgs.Item(lbxFiles.List(i, 0)))
        If chkRefresh Then
        oDwgi.Update
        End If
        Dim OutName As String
        'OutName = Mid(Mid(lbxFiles.List(i, 0), InStrRev(lbxFiles.List(i, 0), "\") + 1), 1, Len(Mid(lbxFiles.List(i, 0), InStrRev(lbxFiles.List(i, 0), "\") + 1)) - 11)
        OutName = Replace(Mid(lbxFiles.List(i, 0), InStrRev(lbxFiles.List(i, 0), "\") + 1), ".", "_")
        If cmbx_Type.Text = "pdf" Then
        Call ExportPDF(oDwgi, oOutFolder.Path & "\" & OutName)
        Else
        oDwgi.ExportData oOutFolder.Path & "\" & OutName, cmbx_Type.Text
        End If
        If Err.Number <> 0 Then
           sResult = "-图纸失败!"
           iFail = iFail + 1
        End If
     On Error GoTo 0
        
     On Error Resume Next
        If chkStp Then
         Dim stpFile
         Set stpFile = DwgLinkedDoc(oDwgi)
         stpFile.ExportData oOutFolder.Path & "\" & Replace(stpFile.Name, ".", "_"), "stp"
         stpFile.Close
        End If
        If Err.Number <> 0 Then
           If sResult <> "-图纸失败!" Then
                iFail = iFail + 1
           End If
           sResult = sResult & "-数模失败!"
        End If
        
     On Error GoTo 0
     If sResult = "" Then
        sResult = "-成功"
     End If
     On Error Resume Next
        lbxFiles.List(i, 0) = lbxFiles.List(i, 0) & sResult
        
    lbxFiles.Selected(i) = False
    oDwgi.Close
    On Error Resume Next
Next

cmdRunBatch.Enabled = False
        If iFail = 0 Then
            MsgBox "操作了 " & lbxFiles.ListCount & " 个图纸文件,全部成功！" _
            & vbCrLf & "文件存放于 " & oOutFolder.Path & "\ 目录"
        Else
            MsgBox "操作了 " & lbxFiles.ListCount & " 个图纸文件, " & iFail & " 个失败！" _
            & vbCrLf & "文件存放于 " & oOutFolder.Path & "\ 目录"
        
        End If
DicDwgs.RemoveAll
CATIA.DisplayFileAlerts = True
CATIA.RefreshDisplay = True

cmdRunBatch.Caption = sCap & "    #输出完成！请重新选择图纸#"
cmdRunBatch.BackColor = &H8000000F

End Sub

Private Sub cmdselFiles_Click()
On Error Resume Next
Dim ofs As Object
    Dim fd As Object
    Set ofs = CreateObject("Excel.Application")

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = ofs.FileDialog(1)
    'fd.Show = True

    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd
         'Add a filter that includes GIF and JPEG images and make it the first item in the list.
        .Filters.Add "DrawingDocument", "*.CATDrawing", 1
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'The user pressed the action button.
        If .Show = -1 Then

            'Step through each string in the FileDialogSelectedItems collection.
            For Each vrtSelectedItem In .SelectedItems

                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
'                MsgBox "The path is: " & vrtSelectedItem & _
'                       vbCrLf & Mid(vrtSelectedItem, InStrRev(vrtSelectedItem, "\") + 1)
'                MsgBox TypeName(vrtSelectedItem)
'lbxFiles.AddItem vrtSelectedItem
'DicDwgs.Add Mid(vrtSelectedItem, InStrRev(vrtSelectedItem, "\") + 1), vrtSelectedItem
'DicDwgs.Add Mid(Mid(vrtSelectedItem, InStrRev(vrtSelectedItem, "\") + 1), 1, Len(Mid(vrtSelectedItem, InStrRev(vrtSelectedItem, "\") + 1)) - 11), vrtSelectedItem
DicDwgs.Add vrtSelectedItem, vrtSelectedItem
            Next vrtSelectedItem
        'The user pressed Cancel.
        Else
        End If
    End With
lbxFiles.List = DicDwgs.keys()
cmdRunBatch.Enabled = True
cmdRunBatch.Caption = sCap
    'Set the object variable to Nothing.
    Set fd = Nothing
On Error GoTo 0
End Sub

Private Sub lblOut1_Click()
On Error Resume Next

Shell "explorer.exe " & lblOut1.Caption, vbNormalFocus
On Error GoTo 0
End Sub


Private Sub UserForm_Activate()

IntCATIA

Set oOutFolder = oCATVBA_Folder("BatchPrint")
lblOut1.Caption = oOutFolder.Path
sCap = cmdRunBatch.Caption



cmbx_Type.AddItem "pdf"
cmbx_Type.AddItem "dxf"
cmbx_Type.AddItem "dwg"
cmbx_Type.AddItem "tif"
cmbx_Type.AddItem "jpg"
cmbx_Type.AddItem "cgm"
cmbx_Type.AddItem "svg"

Set DicDwgs = CreateObject("Scripting.Dictionary") 'CATDrawing list

End Sub


Private Sub ExportPDF(ByVal oDocPDF, ByVal sName)
On Error Resume Next
Dim settingControllers1 'As SettingControllers
Dim settingRepository1 'As SettingRepository
Dim tempB
Set settingControllers1 = CATIA.SettingControllers
Set settingRepository1 = settingControllers1.Item("DraftingOptions")
tempB = settingRepository1.GetAttr("DimDesignMode")
'msgbox tempB
settingRepository1.PutAttr "DimDesignMode", False
'msgbox settingRepository1.GetAttr("DimDesignMode")
settingRepository1.Commit

   oDocPDF.ExportData sName, "pdf"
    
settingRepository1.PutAttr "DimDesignMode", tempB
settingRepository1.Commit
On Error GoTo 0
End Sub

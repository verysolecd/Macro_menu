Public T1, T2, T3
Private Sub CommandButton14_Click()
	On Error Resume Next
	If OptionButton4.Value = False And OptionButton5.Value = False Then
		Exit Sub
	End If
	CATIA.Application.HSOSynchronized = False
	CATIA.Application.RefreshDisplay = False
	CATIA.Application.DisplayFileAlerts = False '--------------批量截图
	Dim productDocument1 As Object
	Set productDocument1 = CATIA.ActiveDocument
	productDocument1.Product.ApplyWorkMode DESIGN_MODE
	Dim product1 As Object
	Set product1 = productDocument1.Application.Documents
	' CATIA.Application.ScreenUpdating = False '关闭屏幕刷新
	'    Application.EnableEvents = False '禁用事件
	Call withe
Dim oPath
If Path.Text = "" Then
oPath = product1.Item(1).Path
Else
oPath = Path.Text
End If
Set oFileSystem = CATIA.FileSystem
Dim FoldObj As Folder
Dim Exists As Boolean
'OPath = InputBox("输入文件夹路径", "批量截图", "") & "\"
Exists = oFileSystem.FolderExists(oPath & "\Pictures")
If Exists = False Then
Set FoldObj = oFileSystem.CreateFolder(oPath & "\Pictures")
End If
    If OptionButton3.Value = True Then
        Set xlApp = CreateObject("excel.application")
        xlApp.Visible = True
        F_Name = product1.Item(1).Product.PartNumber & "_属性清单.xlsx"
        xlApp.Workbooks.Open (oPath & "\" & F_Name)
            If Err.Number = 0 Then
            Else
            MsgBox "无清单"
            Exit Sub
            End If
        Set Book = xlApp.Workbooks
        Set xlSheet1 = xlApp.Workbooks.Item(1).Sheets(1)
        '----------------------------------------------------
        aName = CATIA.ActiveDocument.Product.Name
        Count2 = xlApp.CountA(xlSheet1.Columns(1))
        Dim rngs
        For Each rngs In xlSheet1.Range("B2:B" & Count2)
                If rngs <> "" Then
                    CATIA.Documents.Open (product1.Item(rngs & ".CATProduct").FullName)
                    If Err.Number <> 0 Then
                        Err.Clear
                        CATIA.Documents.Open (product1.Item(rngs & ".CATPart").FullName)
                    End If
                    Set Viewer = CATIA.ActiveWindow.ActiveViewer
                    Viewer.Viewpoint3D = CATIA.ActiveDocument.Cameras.Item(1).Viewpoint3D
                    Viewer.Viewpoint3D.Application.ActiveWindow.Height = 600
                    Viewer.Viewpoint3D.Application.ActiveWindow.Width = 600
                    If OptionButton4.Value = True Then
                        CATIA.StartCommand "规格"
                        CATIA.StartCommand "指南针"
                    End If
                    If OptionButton5.Value = True Then
                        CATIA.StartCommand "Specifications"
                        CATIA.StartCommand "Compass"
                    End If
                    oName = CATIA.ActiveDocument.Product.Name
                    Viewer.Reframe
                    Viewer.CaptureToFile 5, oPath & "\Pictures\" & oName & ".jpg"
                    If OptionButton4.Value = True Then
                        CATIA.StartCommand "规格"
                        CATIA.StartCommand "指南针"
                    End If
                    If OptionButton5.Value = True Then
                        CATIA.StartCommand "Specifications"
                        CATIA.StartCommand "Compass"
                    End If
                        If rngs <> aName Then
                        CATIA.ActiveDocument.Close
                        End If
                End If
        Next
        mr
        MsgBox "已存放至" & oPath & "\Pictures\文件夹中", , "批量截图"
        Exit Sub
        '-----------------------------------------
    Else
        T1 = Split(TextBox2.Text, ",")(0)
        T2 = Split(TextBox2.Text, ",")(1)
        T3 = Split(TextBox2.Text, ",")(2)
        If T1 = "" Then
        T1 = " "
        End If
        If T2 = "" Then
        T2 = " "
        End If
        If T3 = "" Then
        T3 = " "
        End If
    End If
Count1 = product1.Count
For i = 1 To Count1
        oName = product1.Item(i).Product.Name
        oName1 = product1.Item(i).Name
        If OptionButton1.Value = True Then
        GoTo 100
        ElseIf OptionButton2.Value = True Then
            If oName1 Like "*.CATPart" Then
            GoTo 100
            Else
            GoTo 200
            End If
        End If
100
            If oName Like "*" & T1 & "*" Or oName Like "*" & T2 & "*" Or oName Like "*" & T3 & "*" Then
            Else
            CATIA.Documents.Open (product1.Item(i).FullName)
            Set Viewer = CATIA.ActiveWindow.ActiveViewer
            Viewer.Viewpoint3D = CATIA.ActiveDocument.Cameras.Item(1).Viewpoint3D
            Viewer.Viewpoint3D.Application.ActiveWindow.Height = 600
            Viewer.Viewpoint3D.Application.ActiveWindow.Width = 600
            If OptionButton4.Value = True Then
                CATIA.StartCommand "规格"
                CATIA.StartCommand "指南针"
            End If
            If OptionButton5.Value = True Then
                CATIA.StartCommand "Specifications"
                CATIA.StartCommand "Compass"
            End If
            Viewer.Reframe
            Viewer.CaptureToFile 5, oPath & "\Pictures\" & oName & ".jpg"
            If OptionButton4.Value = True Then
                CATIA.StartCommand "规格"
                CATIA.StartCommand "指南针"
            End If
            If OptionButton5.Value = True Then
                CATIA.StartCommand "Specifications"
                CATIA.StartCommand "Compass"
            End If
                If i <> 1 Then
                CATIA.ActiveDocument.Close
                Else
                End If
            End If
200
Next
Call mr
CATIA.Application.RefreshDisplay = True
CATIA.Application.HSOSynchronized = True
MsgBox "已存放至" & oPath & "\Pictures\文件夹中", , "批量截图"
CATIA.Application.DisplayFileAlerts = True
End Sub
Private Sub CommandButton15_Click()  '取消按钮
	Unload Me
	End Sub
	
	
Private Sub CommandButton16_Click()  '取消按钮
	On Error Resume Next
	Set xlApp = CreateObject("excel.application")
	xlApp.Visible = True
	oPath = CATIA.ActiveDocument.Path
	F_Name = CATIA.ActiveDocument.Product.PartNumber & "_属性清单.xlsx"
	xlApp.Workbooks.Open (oPath & "\" & F_Name)
End Sub


Private Sub CommandButton17_Click()
	If Path.Text = "" Then
		oPath = CATIA.ActiveDocument.Path
	Else
		oPath = Path.Text
	End If
	Shell "explorer.exe " & oPath, vbNormalFocus
End Sub
Private Sub OptionButton1_Click()
	If OptionButton1.Value = True Then
		CommandButton16.Visible = False
		Label3.Visible = False
		Label2.Visible = True
		TextBox2.Visible = True
	End If
End Sub
Private Sub OptionButton2_Click()
	If OptionButton2.Value = True Then
	CommandButton16.Visible = False
	Label3.Visible = False
	Label2.Visible = True
	TextBox2.Visible = True
	End If
End Sub

Private Sub OptionButton3_Click()
		On Error Resume Next
		If OptionButton3.Value = True Then
		CommandButton16.Visible = True
		Label3.Visible = True
		Label2.Visible = False
		TextBox2.Visible = False
		Label3.Caption = CATIA.ActiveDocument.Path
		End If
End Sub

Private Sub Path_Change()
End Sub

Private Sub Path_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
	On Error Resume Next
	Path.Text = CATIA.ActiveDocument.Path
End Sub
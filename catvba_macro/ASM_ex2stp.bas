Attribute VB_Name = "ASM_ex2stp"
'Attribute VB_Name = "m35_ex2stp"
' 导出stp并压缩为压缩包
'{GP:3}
'{EP:ex2stp_zip}
'{Caption:导出stp}
'{ControlTipText: 一键导出stp并压缩到指定路径或本身目录}
'{BackColor:}
' 定义模块级变量
Private errorMessage As String
Private formTitle
Private textbox


Sub ex2stp_zip()
    On Error Resume Next ' 临时开启错误处理
    If Not CanExecute("ProductDocument") Then
       errorMessage = "当前文档类型不支持此操作。"
        GoTo ShowMessage
    End If
    
    Dim odoc As Document
    Set odoc = CATIA.ActiveDocument
    
    Dim outputpath As String
    askdir.Show
    outputpath = GetOutputPath(odoc)
    
    If outputpath = "" Then
        errorMessage = "缺少导出路径，操作取消！"
        GoTo ShowMessage
    Else
        Dim pn: pn = odoc.Product.PartNumber
        If dt_pth_ctrl(0) = 1 Then
            Dim ttp: ttp = KCL.timestamp(dt_pth_ctrl(1))
            odoc.Product.PartNumber = KCL.strbflast(pn, "_") & ttp ' 零件号更新
        End If

        stpname = KCL.strbf1st(odoc.Product.PartNumber, "_") & "_Housing_" & ttp
        
        Dim stpfilepath As String
        Dim opath(2) '0=路径，1=name，2=extname
            opath(0) = outputpath
            opath(1) = stpname
            opath(2) = "stp"
        stpfilepath = KCL.JoinPathName(opath)
        
        MsgBox stpfilepath
        '================导出stp
        odoc.ExportData stpfilepath, "stp"
        
        If Err.Number <> 0 Then
            errorMessage = "STP 导出失败：" & Err.Description
            GoTo ShowMessage
        End If
        '================检查文件存在性
                If Not KCL.isExists(stpfilepath) Then
                    errorMessage = "STP 文件导出后未找到：" & stpfilepath
                    GoTo ShowMessage
                End If
                If Not ex2zip(stpfilepath) Then
                    GoTo ShowMessage
                 End If
        KCL.DeleteMe stpfilepath ' 删除原始 STP 文件
    End If

ShowMessage:
    If errorMessage <> "" Then
        MsgBox errorMessage, vbCritical
    Else
        MsgBox stpfilepath & ".zip文件已压缩,STP 原始文件已删除。", vbInformation
    End If
    
    Set odoc = Nothing
    On Error GoTo 0 ' 关闭错误处理
    errorMessage = "" ' 重置错误信息
End Sub

Function ex2zip(oFilepath) As Boolean
    Dim zipPath, result, shell, cmd, path7z
    path7z = "D:\for use\7-Zip\7z.exe"
    If KCL.isExists(path7z) Then
        zipPath = oFilepath & ".7z" ' 构建 7z 压缩包路径
        cmd = """" & path7z & """ a -t7z -mx=9 """ & zipPath & """ """ & oFilepath & """"
    Else
        zipPath = oFilepath & ".zip" ' 构建 ZIP 压缩包路径
        cmd = "powershell -Command ""Compress-Archive -Path '""" & oFilepath & """' -DestinationPath '""" & zipPath & """' -CompressionLevel Optimal -Force"""
    End If
    Set shell = CreateObject("WScript.Shell")
    result = shell.Run(cmd, 0, True)
    If result <> 0 Then
        errorMessage = "压缩失败！请确保 PowerShell 版本不低于 5.0 或 7-Zip 已经安装。"
        ex2zip = False
    Else
        If Not KCL.isExists(zipPath) Then
            errorMessage = "压缩完成但未找到压缩文件！"
            ex2zip = False
        Else
            ex2zip = True
            cmd = "explorer.exe /select, """ & zipPath & """"
        shell.Run (cmd)
        End If
    End If
End Function


Private Function GetOutputPath(ByVal doc As Document) As String
    Select Case dt_pth_ctrl(2)
        Case 0  ' 用户选择自定义路径
            Dim shellApp, folderBrowser
            Set shellApp = CreateObject("Shell.Application")
            Set folderBrowser = shellApp.BrowseForFolder(0, "选择STP输出文件夹", 16, 0)
            If Not folderBrowser Is Nothing Then
                GetOutputPath = folderBrowser.Self.path
            Else
              GetOutputPath = ""
            End If
        Case 1  ' 使用当前文档路径
            GetOutputPath = IIf(doc.path = "", "", doc.path)
        Case others ' 用户取消操作
            GetOutputPath = ""
    End Select
End Function



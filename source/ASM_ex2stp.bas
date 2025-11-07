Attribute VB_Name = "ASM_ex2stp"
'Attribute VB_Name = "m35_ex2stp"
' 导出stp并压缩为压缩包
'{GP:3}
'{EP:ex2stp_zip}
'{Caption:导出stp}
'{ControlTipText: 一键导出stp并压缩到指定路径或本身目录}
'{BackColor:}
' 定义模块级变量
Private ErrorMessage As String
Private formTitle
Private textbox

Sub ex2stp_zip()
    If Not CanExecute("ProductDocument") Then Exit Sub
  On Error Resume Next ' 临时开启错误处理
  Err.Number = 0
    Dim odoc As Document
    Set odoc = CATIA.ActiveDocument
    Dim outputpath As String
    askdir.Show
    outputpath = GetOutputPath(odoc)
    If outputpath = "" Then
        ErrorMessage = "缺少导出路径，操作取消！"
        GoTo ShowMessage
    Else
    '==================零件号和路径相关
        Dim pn: pn = odoc.Product.PartNumber
        Dim ttp: ttp = KCL.timestamp(export_CFG(1))
        If export_CFG(0) = 1 Then  '若更新时间戳' 零件号更新
            odoc.Product.PartNumber = KCL.strbflast(pn, "_") & ttp
        End If
    
        stpname = KCL.strbf1st(odoc.Product.PartNumber, "_") & "_" & ttp
        Dim stpfilepath As String
        Dim opath(2) '0=路径，1=name，2=extname
            opath(0) = outputpath
            opath(1) = stpname
            opath(2) = "stp"
        stpfilepath = KCL.JoinPathName(opath)
     '================导出stp
        odoc.ExportData stpfilepath, "stp"
         '================检查文件存在性
        If Not KCL.isExists(stpfilepath) Then
            ErrorMessage = "STP 文件导出后未找到：" & stpfilepath
            GoTo ShowMessage
        Else
            '============生成导出日志
            If export_CFG(3) <> "" Then
                logpath = opath(0) & "\" & "stp_export_log.md"
                loginfo = "## " & KCL.timestamp("day") & "  " & stpname & ".stp" & vbCrLf & _
                        "  " & export_CFG(3)
                KCL.Appendtext KCL.getmd(logpath), loginfo
            End If
            If Not ex2zip(stpfilepath) Then
                GoTo ShowMessage
            End If
            If Err.Number <> 0 Then
                ErrorMessage = "STP 导出失败：" & Err.Description
                GoTo ShowMessage
            End If
        End If
    End If
ShowMessage:
    If ErrorMessage <> "" Then
        MsgBox ErrorMessage, vbCritical
    Else
        MsgBox "导出成功" & vbCrLf & _
                stpfilepath & "文件已压缩" & vbCrLf & _
                "原始文件已删除。", vbInformation
    End If
    Set odoc = Nothing
    On Error GoTo 0 ' 关闭错误处理
    ErrorMessage = "" ' 重置错误信息
    
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
    ex2zip = False
    
    If result <> 0 Then
        ErrorMessage = "压缩失败！请确保 PowerShell 版本不低于 5.0 或 7-Zip 已经安装。"
        Err.Clear
    End If
    If KCL.isExists(zipPath) Then
            ex2zip = True
            KCL.DeleteMe oFilepath ' 删除原始 STP 文件
        cmd = "explorer.exe /select, """ & zipPath & """"
        shell.Run (cmd)
    End If

End Function

Private Function GetOutputPath(ByVal doc As Document) As String
    Select Case export_CFG(2)
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



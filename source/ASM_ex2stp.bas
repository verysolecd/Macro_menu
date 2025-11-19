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
Private zippath

Sub ex2stp_zip()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
  On Error Resume Next ' 临时开启错误处理
    Err.Number = 0
    ErrorMessage = ""
    Dim odoc: Set odoc = CATIA.ActiveDocument
    Dim outputpath As String: outputpath = ""
    askdir.Show
    Select Case export_CFG(2)
        Case 0  ' 用户选择自定义路径
              outputpath = KCL.selFdl()
        Case 1  ' 使用当前文档路径
            outputpath = IIf(odoc.path = "", "", odoc.path)
        Case others ' 用户取消操作
            outputpath = ""
    End Select
        If outputpath = "" Then
            ErrorMessage = "缺少导出路径，操作取消！"
            GoTo ShowMessage
        End If
        '==================零件号和路径
     
        
        Dim ttp: ttp = KCL.timestamp(export_CFG(1))
         Dim pn
            If export_CFG(0) = 1 Then  '若更新时间戳 ' 零件号更新
                pn = KCL.strbflast(odoc.Product.PartNumber, "_")
                If KCL.ExistsKey(pn, "_") Then
                    odoc.Product.PartNumber = pn & ttp
                Else
                    odoc.Product.PartNumber = pn & "_" & ttp
                End If
            End If
            pn = odoc.Product.PartNumber
        
            stpname = KCL.strbf1st(pn, "_") & "_" & ttp
            
        Dim oPath(2) '0=路径，1=name，2=extname
            oPath(0) = outputpath
            oPath(1) = stpname
            oPath(2) = "stp"
       Dim stpfilepath As String: stpfilepath = KCL.JoinPathName(oPath)
       
        odoc.ExportData stpfilepath, "stp"     '=======导出stp
        If Not KCL.isExists(stpfilepath) Then '=======检查文件存在性
            ErrorMessage = "未找到STP文件：" & stpfilepath
            GoTo ShowMessage
        End If
        
        If Not ex2zip(stpfilepath) Then GoTo ShowMessage
            KCL.DeleteMe stpfilepath ' 删除原始 STP 文件
        
        If export_CFG(3) <> "" Then '============生成导出日志
            logpath = oPath(0) & "\" & "stp_export_log.md"
            loginfo = "## " & KCL.timestamp("day") & "  " & stpname & ".stp" & vbCrLf & _
                    "  " & export_CFG(3)
            KCL.Appendtext KCL.getmd(logpath), loginfo
        End If
        If Err.Number <> 0 Then
                ErrorMessage = "STP 导出失败：" & Err.Description
                GoTo ShowMessage
        End If
ShowMessage:
    If ErrorMessage <> "" Then
        MsgBox ErrorMessage, vbCritical
        Err.Clear
        Exit Sub
    Else
        MsgBox "导出成功" & vbCrLf & _
                stpfilepath & "文件已压缩" & vbCrLf & _
                "原始文件已删除。", vbInformation
        KCL.openpath (zippath)
    End If
    Set odoc = Nothing
    On Error GoTo 0 ' 关闭错误处理
    ErrorMessage = "" ' 重置错误信息
End Sub

Function ex2zip(oFilepath) As Boolean
 On Error GoTo seterror
    ex2zip = False
    Dim result, shell, cmd, path7z
    '=======================
        path7z = "D:\for use\7-Zip\7z.exe"
        If KCL.isExists(path7z) Then
            zippath = oFilepath & ".7z" ' 构建 7z 压缩包路径
            cmd = """" & path7z & """ a -t7z -mx=9 """ & zippath & """ """ & oFilepath & """"
        Else
            zippath = oFilepath & ".zip" ' 构建 ZIP 压缩包路径
            cmd = "powershell -Command ""Compress-Archive -Path '""" & oFilepath & """' -DestinationPath '""" & zippath & """' -CompressionLevel Optimal -Force"""
        End If
    '=======================
    Set shell = CreateObject("WScript.Shell")
    result = shell.Run(cmd, 0, True)
    If result <> 0 Then
        Err.Clear
    End If
    If KCL.isExists(zippath) Then
            ex2zip = True
            Exit Function
    End If
seterror:
        ErrorMessage = "压缩失败！请确保 PowerShell 版本不低于 5.0 或 7-Zip 已经安装。"
        ex2zip = fasle
    On Error GoTo 0
    
End Function



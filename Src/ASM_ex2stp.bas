Attribute VB_Name = "ASM_ex2stp"

'------宏信息-----------------------------------------------------
'{GP:3}
'{EP:ex2stp_zip}
'{Caption:导出stp}
'{ControlTipText: 一键导出stp并压缩到指定路径或本身目录}
'{BackColor:}

'------窗体标题-------------------------------------------------
'标题格式为 %Title <Caption/Text>
'%Title 现在要导出stp,那请问你?

'------控件清单--------------------------------------------------

'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_path  是否导出到当前路径
' %UI CheckBox  chk_tm  是否更新时间戳到CATIA零件号？
' %UI CheckBox chk_log  是否更新本次导出日志？
' %UI TextBox   txt_log  请输入更新内容信息,不必输入日期
' %UI Button btnOK  确定
' %UI Button btncancel  取消
'------------------------------------------------

Private ErrorMessage As String
Private zippath

Sub ex2stp_zip()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
  On Error Resume Next ' 临时开启错误处理
    Err.Number = 0
    ErrorMessage = ""
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Dim outputpath As String: outputpath = ""
    
    Dim frmDic: Set frmDic = getFrmDic ' oFrm.Res
    
    If frmDic("btn_clicked") <> "btnOK" Then   ' 2. 检查是否点击了确定 (btnOK)
        MsgBox "用户取消了操作"
        Exit Sub
    End If
    '===========路径设置
    If frmDic.Exists("chk_path") = True Then
               Select Case frmDic("chk_path")
                    Case True: outputpath = IIf(oDoc.path = "", "", oDoc.path)
                    Case fasle: outputpath = KCL.selFdl()
                End Select
     End If
    If outputpath = "" Then
            ErrorMessage = "缺少导出路径，操作取消！"
            GoTo ShowMessage
        End If
    '===========零件号时间戳处理
      If frmDic.Exists("chk_tm") = True Then
           Dim ttp: ttp = KCL.timestamp("min")
               Select Case frmDic("chk_tm")
                    Case True:
                           pn = KCL.strbflast(oDoc.Product.PartNumber, "_")
                                   If KCL.ExistsKey(pn, "_") Then
                                       oDoc.Product.PartNumber = pn & ttp
                                   Else
                                       oDoc.Product.PartNumber = pn & "_" & ttp
                                   End If
                    Case fasle:
                End Select
      End If
          pn = oDoc.Product.PartNumber
          
        '==========STP文件名处理
          stpname = KCL.strbf1st(pn, "_") & "_" & ttp
        Dim opath(2) '0=路径，1=name，2=extname
            opath(0) = outputpath
            opath(1) = stpname
            opath(2) = "stp"
       Dim stpfilepath As String: stpfilepath = KCL.JoinPathName(opath)
        oDoc.ExportData stpfilepath, "stp"     '=======导出stp
        If Not KCL.isExists(stpfilepath) Then '=======检查文件存在性
            ErrorMessage = "未找到STP文件：" & stpfilepath
            GoTo ShowMessage
        End If
        
        If Not ex2zip(stpfilepath) Then GoTo ShowMessage
            KCL.DeleteMe stpfilepath ' 删除原始 STP 文件
'============生成导出日志
  If frmDic.Exists("chk_log") Then
     Select Case frmDic("chk_log")
          Case True:
               logpath = opath(0) & "\" & "stp_export_log.md"
               loginfo = "## " & KCL.timestamp("day") & "  " & stpname & ".stp" & vbCrLf & _
                         "  " & frmDic("txt_log")
               KCL.Appendtext KCL.getmd(logpath), loginfo
          Case False:
     End Select
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
    Set oDoc = Nothing
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

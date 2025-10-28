Attribute VB_Name = "ASM_ex2stp"
'Attribute VB_Name = "m35_ex2stp"
' ����stp��ѹ��Ϊѹ����
'{GP:3}
'{EP:ex2stp_zip}
'{Caption:����stp}
'{ControlTipText: һ������stp��ѹ����ָ��·������Ŀ¼}
'{BackColor:}
' ����ģ�鼶����
Private errorMessage As String

Sub ex2stp_zip()
    On Error Resume Next ' ��ʱ����������
    If Not CanExecute("ProductDocument") Then
       errorMessage = "��ǰ�ĵ����Ͳ�֧�ִ˲�����"
        GoTo ShowMessage
    End If
    
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    Dim outputpath As String
    askdir.Show
    outputpath = GetOutputPath(oDoc)
    
    If outputpath = "" Then
        errorMessage = "ȱ�ٵ���·��������ȡ����"
        GoTo ShowMessage
    Else
        Dim pn: pn = oDoc.Product.PartNumber
        If dt_pth_ctrl(0) = 1 Then
            Dim ttp: ttp = KCL.timestamp(dt_pth_ctrl(1))
            oDoc.Product.PartNumber = KCL.strbflast(pn, "_") & ttp ' ����Ÿ���
        End If

        stpname = KCL.strbf1st(oDoc.Product.PartNumber, "_") & "_Housing_" & ttp
        
        Dim stpfilepath As String
        Dim opath(2) '0=·����1=name��2=extname
            opath(0) = outputpath
            opath(1) = stpname
            opath(2) = "stp"
        stpfilepath = KCL.JoinPathName(opath)
        
        MsgBox stpfilepath
        '================����stp
        oDoc.ExportData stpfilepath, "stp"
        If Err.Number <> 0 Then
            errorMessage = "STP ����ʧ�ܣ�" & Err.Description
            GoTo ShowMessage
        End If
        '================����ļ�������
                If Not KCL.isExists(stpfilepath) Then
                    errorMessage = "STP �ļ�������δ�ҵ���" & stpfilepath
                    GoTo ShowMessage
                End If
                If Not ex2zip(stpfilepath) Then
                    GoTo ShowMessage
                 End If
        KCL.DeleteMe stpfilepath ' ɾ��ԭʼ STP �ļ�
    End If

ShowMessage:
    If errorMessage <> "" Then
        MsgBox errorMessage, vbCritical
    Else
        MsgBox stpfilepath & ".zip�ļ���ѹ��,STP ԭʼ�ļ���ɾ����", vbInformation
    End If
    
    Set oDoc = Nothing
    On Error GoTo 0 ' �رմ�����
    errorMessage = "" ' ���ô�����Ϣ
End Sub

Function ex2zip(oFilepath) As Boolean
    Dim zipPath, result, shell, cmd, path7z
    path7z = "D:\for use\7-Zip\7z.exe"
    If KCL.isExists(path7z) Then
        zipPath = oFilepath & ".7z" ' ���� 7z ѹ����·��
        cmd = """" & path7z & """ a -t7z -mx=9 """ & zipPath & """ """ & oFilepath & """"
    Else
        zipPath = oFilepath & ".zip" ' ���� ZIP ѹ����·��
        cmd = "powershell -Command ""Compress-Archive -Path '""" & oFilepath & """' -DestinationPath '""" & zipPath & """' -CompressionLevel Optimal -Force"""
    End If
    Set shell = CreateObject("WScript.Shell")
    result = shell.Run(cmd, 0, True)
    If result <> 0 Then
        errorMessage = "ѹ��ʧ�ܣ���ȷ�� PowerShell �汾������ 5.0 �� 7-Zip �Ѿ���װ��"
        ex2zip = False
    Else
        If Not KCL.isExists(zipPath) Then
            errorMessage = "ѹ����ɵ�δ�ҵ�ѹ���ļ���"
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
        Case 0  ' �û�ѡ���Զ���·��
            Dim shellApp, folderBrowser
            Set shellApp = CreateObject("Shell.Application")
            Set folderBrowser = shellApp.BrowseForFolder(0, "ѡ��STP����ļ���", 16, 0)
            If Not folderBrowser Is Nothing Then
                GetOutputPath = folderBrowser.Self.path
            Else
              GetOutputPath = ""
            End If
        Case 1  ' ʹ�õ�ǰ�ĵ�·��
            GetOutputPath = IIf(doc.path = "", "", doc.path)
        Case others ' �û�ȡ������
            GetOutputPath = ""
    End Select
End Function



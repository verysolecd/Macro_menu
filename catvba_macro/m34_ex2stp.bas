Attribute VB_Name = "m34_ex2stp"
'Attribute VB_Name = "m34_ex2stp"
' ����stp��ѹ��Ϊѹ����
'{GP:3}
'{EP:ex2stp_zip}
'{Caption:����stp}
'{ControlTipText: һ������stp��ѹ����ָ��·������Ŀ¼}
'{BackColor:12648447}

' ����ģ�鼶����
Private errorMessage As String

Sub ex2stp_zip()
    On Error Resume Next ' ��ʱ����������
    If Not CanExecute("ProductDocument") Then
        errorMessage = "��ǰ�ĵ����Ͳ�֧�ִ˲�����"
        GoTo ShowMessage
    End If
    
    Dim odoc As Document
    Set odoc = CATIA.ActiveDocument
    
    Dim outputpath As String
    outputpath = GetOutputPath(odoc)
    
    If outputpath = "" Then
        errorMessage = "ȱ�ٵ���·��������ȡ����"
        GoTo ShowMessage
    Else
        Dim tdy As String
        tdy = Format(Now, "yymmdd.hh.nn")
        Dim pn As String
        pn = odoc.product.PartNumber
        odoc.product.PartNumber = Pntdy(pn, tdy)  ' ����Ÿ���
        Dim stpfilepath As String
        
        Dim opath(2) '0=·����1=name��2=extname
        opath(0) = outputpath
        opath(1) = GetSTPFileName(odoc.product)
        opath(2) = "stp"
        
        stpfilepath = KCL.JoinPathName(opath)
        MsgBox stpfilepath
    '    stpfilepath = outputpath & "\" & GetSTPFileName(oDoc.product) & ".stp"
        odoc.ExportData stpfilepath, "stp"
        If Err.Number <> 0 Then
            errorMessage = "STP ����ʧ�ܣ�" & Err.Description
            GoTo ShowMessage
        End If
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
    
    Set odoc = Nothing
    On Error GoTo 0 ' �رմ�����
    errorMessage = "" ' ���ô�����Ϣ
End Sub

Private Function GetOutputPath(ByVal doc As Document) As String
    Dim userChoice As VbMsgBoxResult

    userChoice = MsgBox("ѡ�񵼳�·��:" & vbNewLine & _
    "�� �ǡ�      ѡ�񵼳�·��" & vbNewLine & _
    "�� ��      ������Product����·��" & vbNewLine & _
    "��ȡ�� ��   �˳�", _
    vbYesNoCancel + vbExclamation, "����")
    
    Select Case userChoice
        Case vbYes  ' �û�ѡ���Զ���·��
            Dim shellApp As Object
            Set shellApp = CreateObject("Shell.Application")
            Dim folderBrowser As Object
            Set folderBrowser = shellApp.BrowseForFolder(0, "ѡ��STP����ļ���", 16, 0)
            If Not folderBrowser Is Nothing Then
                GetOutputPath = folderBrowser.Self.path
            Else
              GetOutputPath = ""
            End If
        Case vbNo
            ' ʹ�õ�ǰ�ĵ�·��
            GetOutputPath = IIf(doc.path = "", "", doc.path)
        Case vbCancel
            ' �û�ȡ������
            GetOutputPath = ""
    End Select
End Function

Private Function GetSTPFileName(ByVal product As Object) As String   ' ���ɴ�ʱ�����STP�ļ���
    Dim timestamp As String
    timestamp = Format(Now, "yymmdd_hhnn")
    
    GetSTPFileName = getPrefix(product.PartNumber) & "_Prj_Housing_" & timestamp
End Function


Function Pntdy(text, rep)
Dim lastindex
lastindex = InStrRev(text, "_")
If lastindex > 0 Then
        Pntdy = Left(text, lastindex) & rep
    Else
        Pntdy = text
    End If
End Function

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
        End If
    End If
End Function

'@iStr string
'����ַ�����һ��"_"֮ǰ���ַ��򷵻�ԭ�ַ�

Function getPrefix(iStr)

    Dim oPrefix As String
        Dim underscorePos As Long
        underscorePos = InStr(iStr, "_")
        If underscorePos > 0 Then
            oPrefix = Left(iStr, underscorePos - 1)
        Else
           oPrefix = iStr
        End If
End Function
 



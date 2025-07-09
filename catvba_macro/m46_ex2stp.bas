'Attribute VB_Name = "m30_ex2stp"
' ����stp��ѹ��Ϊѹ����
'{GP:3}
'{EP:ex2stp}
'{Caption:����stp}
'{ControlTipText: һ������stp��ָ��·������Ŀ¼}
'{BackColor:12648447}

' ����ģ�鼶����
Dim errorMessage As String

Sub ex2stp()
    On Error Resume Next ' ��ʱ����������
    
    If Not CanExecute("ProductDocument") Then
        errorMessage = "��ǰ�ĵ����Ͳ�֧�ִ˲�����"
        GoTo ShowMessage
    End If
    
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    Dim outputpath As String
    outputpath = GetOutputPath(oDoc)
    
    If outputpath = "" Then
        errorMessage = "δѡ������ļ��У�����ȡ����"
        GoTo ShowMessage
    End If
    
    Dim tdy As String
    tdy = Format(Now, "yymmdd.hh.nn")
    Dim pn As String
    pn = oDoc.Product.PartNumber
    oDoc.Product.PartNumber = Pntdy(pn, tdy)  ' ����Ÿ���
    
    Dim stpfilepath As String
    stpfilepath = outputpath & "\" & GetSTPFileName(oDoc.Product) & ".stp"
    oDoc.ExportData stpfilepath, "stp"
    
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

ShowMessage:
    If errorMessage <> "" Then
        MsgBox errorMessage, vbCritical
    Else
        MsgBox "STP �ļ��ѳɹ�������ѹ����", vbInformation
    End If
    
    Set oDoc = Nothing
    On Error GoTo 0 ' �رմ�����
    errorMessage = "" ' ���ô�����Ϣ
End Sub

Private Function GetOutputPath(ByVal doc As Document) As String
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("ѡ�񵼳�·��", vbYesNoCancel + vbExclamation, "����")
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

Private Function GetSTPFileName(ByVal product As Object) As String
    ' ���ɴ�ʱ�����STP�ļ���
    Dim timestamp As String
    timestamp = Format(Now, "yymmdd_hhnn")
    Dim filePrefix As String
    Dim underscorePos As Long
    underscorePos = InStr(product.PartNumber, "_")
    If underscorePos > 0 Then
        filePrefix = Left(product.PartNumber, underscorePos - 1)
    Else
        filePrefix = product.PartNumber
    End If
    GetSTPFileName = filePrefix & "_Prj_Housing_" & timestamp
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
Function ex2zip(stppath) As Boolean
    Dim zipPath As String
    Dim result As Long
    Dim shell As Object
    Dim cmd As String
    Dim path7z As String
    path7z = "D:\for use\7-Zip\7z.exe"    
    If KCL.isExists(path7z) Then
        zipPath = stppath & ".7z" ' ���� 7z ѹ����·��
        cmd = """" & path7z & """ a -t7z -mx=9 """ & zipPath & """ """ & stppath & """"
    Else
        zipPath = stppath & ".zip" ' ���� ZIP ѹ����·��
        cmd = "powershell -Command ""Compress-Archive -Path '""" & stppath & """' -DestinationPath '""" & zipPath & """' -CompressionLevel Optimal -Force"""
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

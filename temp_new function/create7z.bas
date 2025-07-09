Function AddToMarkdown() As Boolean
    Dim filePath As String
    Dim userInput As String
    Dim fso As Object
    Dim ts As Object
    
    ' ����Markdown�ļ�·�����ɸ�����Ҫ�޸�
    filePath = "C:\Users\YourName\Documents\�ʼ�.md"
    
    ' ��ȡ�û�����
    userInput = InputBox("������Ҫ��ӵ�Markdown�ļ������ݣ�", "�������")
    
    ' ����û��Ƿ�ȡ��������
    If userInput = "" Then
        AddToMarkdown = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' �����ļ�ϵͳ����
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ��׷��ģʽ���ļ������ļ��������򴴽�
    Set ts = fso.OpenTextFile(filePath, 8, True)
    
    ' д���û��������ݣ�����ӻ��з�
    ts.WriteLine userInput
    
    ' �ر��ļ�
    ts.Close
    
    AddToMarkdown = True
    Exit Function
    
ErrorHandler:
    MsgBox "д���ļ�ʱ����: " & Err.Description, vbExclamation
    AddToMarkdown = False
    If Not ts Is Nothing Then ts.Close
End Function



Sub ExportSTPAndCompress()
    On Error Resume Next
    
    ' ��ȡ��ǰ��ĵ�
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' ����Ƿ�Ϊ��Ʒ�ĵ�
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "��ȷ����ǰ�򿪵��ǲ�Ʒ�ĵ�!", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ��Ʒ���ƣ�ȥ����չ����
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' ��ȡ��һ���»���ǰ��ǰ׺����DX11_DDD �� DX11��
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' ��ȡ��ǰ���ڣ���ʽ��ΪYYMMDD��
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' ѡ������ļ���
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "ѡ��STP����ļ���", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "δѡ������ļ��У�����ȡ��!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' ����STP�ļ��������磺DX11_231005.stp��
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' ����ΪSTP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP����ʧ�ܣ�" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' ����7-Zipѹ����·�������磺DX11_231005.7z��
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".7z"
    
    ' ����7-Zip�����й��߽���ѹ��
    Dim shell, cmd
    Set shell = CreateObject("WScript.Shell")
    
    ' 7-Zip���a = ��ӵ�ѹ������-t7z = ָ��7z��ʽ��-mx=9 = ���ѹ����
    cmd = """C:\Program Files\7-Zip\7z.exe"" a -t7z -mx=9 """ & zipPath & """ """ & stpPath & """"
    
    ' ִ������
    Dim result
    result = shell.Run(cmd, 0, True)  ' 0 = ���ش��ڣ�True = �ȴ�����ִ�����
    
    If result <> 0 Then
        MsgBox "7-Zipѹ��ʧ�ܣ���ȷ��7-Zip����ȷ��װ��", vbCritical
    Else
        MsgBox "STP������ѹ���ɹ���" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        
        ' ��ѡ��ɾ��ԭʼSTP�ļ�������ѹ������
        ' If FileExists(stpPath) Then Kill stpPath
    End If
    
    ' �ͷŶ���
    Set shell = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub

' ��������������ļ��Ƿ����
Function FileExists(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function






Sub ExportSTPAndZipWithWin11()
    On Error Resume Next
    
    ' ��ȡ��ǰ��ĵ�
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' ����Ƿ�Ϊ��Ʒ�ĵ�
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "��ȷ����ǰ�򿪵��ǲ�Ʒ�ĵ�!", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ��Ʒ���ƣ�ȥ����չ����
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' ��ȡ��һ���»���ǰ��ǰ׺����DX11_DDD �� DX11��
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' ��ȡ��ǰ���ڣ���ʽ��ΪYYMMDD��
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' ѡ������ļ���
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "ѡ��STP����ļ���", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "δѡ������ļ��У�����ȡ��!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' ����STP�ļ��������磺DX11_231005.stp��
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' ����ΪSTP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP����ʧ�ܣ�" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' ����ZIPѹ����·�������磺DX11_231005.zip��
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".zip"
    
    ' ������ZIP�ļ���Windows 11ԭ��֧�֣�
    CreateEmptyZipFile zipPath
    
    ' �ȴ�ZIP�ļ��������
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim waitTime As Integer
    waitTime = 0
    
    Do While Not fso.FileExists(zipPath) And waitTime < 10
        WScript.Sleep 500  ' �ȴ�0.5��
        waitTime = waitTime + 1
    Loop
    
    If Not fso.FileExists(zipPath) Then
        MsgBox "����ZIP�ļ�ʧ�ܣ�", vbCritical
        Exit Sub
    End If
    
    ' ��STP�ļ���ӵ�ZIPѹ����
    Dim sourceFile, destinationZip
    Set sourceFile = ShellApp.NameSpace(stpPath).Items
    Set destinationZip = ShellApp.NameSpace(zipPath)
    
    If Not destinationZip Is Nothing Then
        destinationZip.CopyHere sourceFile, 4  ' 4 = ����ʾȷ�϶Ի���
        
        ' �ȴ�ѹ����ɣ���������
        WScript.Sleep 2000  ' �ȴ�2�루�ɸ����ļ���С������
        
        MsgBox "STP������ѹ���ɹ���" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        
        ' ��ѡ��ɾ��ԭʼSTP�ļ�������ѹ������
        ' If fso.FileExists(stpPath) Then fso.DeleteFile stpPath
    Else
        MsgBox "�޷�����ZIP�ļ���", vbCritical
    End If
    
    ' �ͷŶ���
    Set fso = Nothing
    Set destinationZip = Nothing
    Set sourceFile = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub

' ������ZIP�ļ��ĸ�������
Sub CreateEmptyZipFile(zipFilePath)
    Dim fso, tempFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ������ʱ�ļ�д��ZIP�ļ�ͷ
    Set tempFile = fso.CreateTextFile(zipFilePath, True)
    tempFile.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
    tempFile.Close
    
    Set tempFile = Nothing
    Set fso = Nothing
End Sub

Sub ExportSTPAndZipWithPowerShell()
    On Error Resume Next
    
    ' ��ȡ��ǰ��ĵ�
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' ����Ƿ�Ϊ��Ʒ�ĵ�
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "��ȷ����ǰ�򿪵��ǲ�Ʒ�ĵ�!", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ��Ʒ���ƣ�ȥ����չ����
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' ��ȡ��һ���»���ǰ��ǰ׺����DX11_DDD �� DX11��
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' ��ȡ��ǰ���ڣ���ʽ��ΪYYMMDD��
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' ѡ������ļ���
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "ѡ��STP����ļ���", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "δѡ������ļ��У�����ȡ��!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' ����STP�ļ��������磺DX11_231005.stp��
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' ����ΪSTP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP����ʧ�ܣ�" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' ����ZIPѹ����·�������磺DX11_231005.zip��
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".zip"
    
    ' ʹ��PowerShell����ѹ���ļ�
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")
    
    ' PowerShell���ʹ�����ѹ������(-CompressionLevel Fastest)
    cmd = "powershell -Command ""Compress-Archive -Path '""" & stpPath & """' -DestinationPath '""" & zipPath & """' -CompressionLevel Fastest -Force"""
    
    ' ִ�������ȡ����ֵ
    result = shell.Run(cmd, 0, True)
    
    If result <> 0 Then
        MsgBox "PowerShellѹ��ʧ�ܣ���ȷ��PowerShell�汾������5.0��", vbCritical
    Else
        ' ��֤ZIP�ļ��Ƿ����
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FileExists(zipPath) Then
            MsgBox "STP������ѹ���ɹ���" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        Else
            MsgBox "ѹ����ɵ�δ�ҵ�ZIP�ļ���", vbCritical
        End If
        
        Set fso = Nothing
    End If
    
    ' �ͷŶ���
    Set shell = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub




Dim currentDateTime As String
currentDateTime = Year(Date) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2) & "_" & _
                  Right("0" & Hour(Now), 2) & _
                  Right("0" & Minute(Now), 2) & _
                  Right("0" & Second(Now), 2)
' �����磺20231001_143045




��ȡ�ļ���
Dim fileName
fileName = "DX11_DDD_20231005.stp"

' ���ҵ�һ���»��ߵ�λ��
Dim underscorePos
underscorePos = InStr(fileName, "_")

' ����ҵ��»��ߣ����ȡǰ����ַ�
Dim prefix
If underscorePos > 0 Then
    prefix = Left(fileName, underscorePos - 1)  ' ��ȡ����߿�ʼ���»���ǰ���ַ�
Else
    prefix = fileName  ' ���û���»��ߣ�����ԭ�ļ���
End If

MsgBox "��ȡ���: " & prefix  ' ���: DX11




Sub ExportSTPWithPrefixAndDate()
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
        prefix = productName  ' ���û���»��ߣ�ʹ����������
    End If
    
    ' ��ȡ��ǰ���ڣ����ȡ����λ����ʽ��ΪYYMMDD��
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
    
    ' ������ǰ׺�����ڵ�STP�ļ��������磺DX11_231005.stp��
    Dim outputPath As String
    outputPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' ����ΪSTP
    oDoc.ExportData outputPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "����ʧ�ܣ�" & Err.Description, vbCritical
    Else
        MsgBox "STP�����ɹ���" & outputPath, vbInformation
    End If
    
    ' �ͷŶ���
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub
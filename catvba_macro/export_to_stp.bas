Attribute VB_Name = "export_to_stp"
' ������ǰ��ĵ�ΪSTP�ļ�
Option Explicit

Sub ExportToSTP()
    ' ����Ƿ����ִ�в���
    If Not KCL.CanExecute("PartDocument,ProductDocument") Then Exit Sub
    
    Dim doc As Document
    Set doc = CATIA.ActiveDocument
    
    ' ���û�ѡ�񱣴�·�����ļ���
    Dim filePath As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "��ѡ�񱣴�STP�ļ���λ��"
        .InitialFileName = "example.stp"
        .Filters.Clear
        .Filters.Add "STEP �ļ�", "*.stp"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            If Right(filePath, 4) <> ".stp" Then
                filePath = filePath & ".stp"
            End If
        Else
            MsgBox "δѡ�񱣴�·��������ȡ����", vbExclamation
            Exit Sub
        End If
    End With
    
    If filePath = "" Then
        MsgBox "δ������Ч�ı���·��������ȡ����", vbExclamation
        Exit Sub
    End If
    
    ' ����ΪSTP�ļ�
    On Error Resume Next
    doc.ExportData filePath, "stp"
    If Err.Number <> 0 Then
        MsgBox "����ʧ�ܣ�" & Err.Description, vbCritical
    Else
        MsgBox "�ļ��ѳɹ���������" & filePath, vbInformation
    End If
    On Error GoTo 0
End Sub
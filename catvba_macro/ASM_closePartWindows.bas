Attribute VB_Name = "ASM_closePartWindows"
'Attribute VB_Name = "m10_closePartWindows"
' Part����һ����ȫ�ر�
'{GP:3}
'{EP:CLSpart}
'{Caption:�ر������}
'{ControlTipText: �����һ����ȫ�ر������������}
'{������ɫ: 12648447}

Sub CLSpart()
Dim CATIA, wds, wd
 On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    Set wds = CATIA.Windows
    If wds.Count < 1 Then
           MsgBox "û�д򿪵Ĵ���"
           Exit Sub
    End If
    For i = 1 To wds.Count
        Set wd = wds.item(i)
        If KCL.IsType_Of_T(wd.Parent, "PartDocument") Then
            wd.Close
        End If
    Next
    On Error GoTo 0
End Sub

'

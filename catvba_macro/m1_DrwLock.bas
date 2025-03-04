Attribute VB_Name = "m1_DrwLock"
'Attribute VB_Name = "sample_Draft_View_Lock_UnLock"
' ͼֽ��ͼ�����������
'{GP:1}
'{EP:CATMain}
'{Caption:����_����}
'{ControlTipText: ���Խ���ͼֽ��ͼ�����������}
'{������ɫ: 12648447}

Option Explicit
Sub CATMain()
' ����Ƿ����ִ��
     If Not CanExecute("DrawingDocument") Then
          Exit Sub
     End If
     
     Dim Views As DrawingViews
     Set Views = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
     If Views.Count < 3 Then
                 Exit Sub
      End If
            
            Dim View As DrawingView
                 Set View = Views.Item(3)
            Dim LockState As Boolean
                 LockState = View.LockStatus
            Dim Msg As String
            
            If LockState Then
                 Msg = "����"
               LockState = False
            Else
                 Msg = "����"
               LockState = True
            End If
     If Views.Count > 3 Then
            Dim i As Long
            For i = 3 To Views.Count
                 Set View = Views.Item(i)
                      View.LockStatus = LockState
                 Next
     End If
     MsgBox "��ͼ�ѳɹ�" & Msg & "��"
End Sub

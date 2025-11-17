Attribute VB_Name = "DRW_DrwLock"
'Attribute VB_Name = "sample_Draft_View_Lock_UnLock"
' 图纸视图的锁定与解锁
'{GP:5}
'{EP:CATMain}
'{Caption:锁定_解锁}
'{ControlTipText: 可以进行图纸视图的锁定与解锁}
'{背景颜色: 12648447}

Option Explicit
Sub CATMain()
' 检查是否可以执行
     If Not CanExecute("DrawingDocument") Then
          Exit Sub
     End If
     
     Dim Views As DrawingViews
     Set Views = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
     If Views.count < 3 Then
                 Exit Sub
      End If
            
            Dim View As DrawingView
                 Set View = Views.item(3)
                 
            Dim LockState As Boolean
                 LockState = View.LockStatus
            
            Dim msg As String
            
            If LockState Then
                 msg = "解锁"
               LockState = False
            Else
                 msg = "锁定"
               LockState = True
            End If
     If Views.count > 3 Then
            Dim i As Long
            For i = 3 To Views.count
                 Set View = Views.item(i)
                      View.LockStatus = LockState
                 Next
     End If
     MsgBox "视图已成功" & msg & "。"
End Sub

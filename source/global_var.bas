Attribute VB_Name = "global_var"
Public gPrd As Object
Public rootDoc
Public rootprd  As Object  '全局产品obj
Public startrow, lastrow  '全局excel行定义
Public xlAPP As Object  '全局excelcom组件
Public gwb As Object
Public gws  As Object
Public pdm As Object
Public xlm As Object
Public allPN As Object
Public counter As Integer
Public Const gfn As Integer = 400
Public ProductObserver As New ProductObserver
Public export_CFG   ' 被定义为一个数组  Ary()  第一个元素是开
Public gPic_Path


Sub clearall()

End Sub

'Dim btn, bTitle, bResult
'imsg = "将备份到" & bckpath "您确认吗"

'btn = vbYesNo + vbExclamation
'bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)

'Select Case bResult
'Case 7: Exit Sub '===选择“否”====
'Case 6  '===选择“是”====
'Case 2  '===选择“取消”====


Public Function showdict(ByVal oDic, Optional ByVal boolShowKeyIndex As Boolean = False)
  Dim keys:   keys = oDic.keys
  Dim i As Long
  Dim stIndex As String
  Dim stOutput As String
  stOutput = vbNullString
  
  For i = 0 To oDic.count - 1
    If boolShowKeyIndex Then
      stIndex = "(" & i & ")"
    End If
    stOutput = stOutput & keys(i) & stIndex & "  :  "
    If IsObject(oDic(keys(i))) Then
      stOutput = stOutput & "[" & showdict(oDic(keys(i)), boolShowKeyIndex) & "]"
    Else
      stOutput = stOutput & oDic(keys(i))
    End If
    stOutput = stOutput & "; " & "_" & vbNewLine
  Next i
  showdict = stOutput
  
  Debug.Print showdict
End Function

Attribute VB_Name = "TEST"
' %UI Button btn_ok 你好
' %UI Button btn_cancel 请输入去哪吃
' %UI chk chk_eatshit  是否

Private btns, ctrls
Private WD

Sub test()
    Set WD = New Cls_DynaFrm: If WD.IsCancelled Then Exit Sub
'-----------------按钮的使用-----------------
btns = Array("btn_nihao", "", "", "")
'    Select Case WD.BtnClicked
'        Case btns(0): MsgBox "决定吃屎"
'        Case btns(1):
'        Case btns(2):
'        Case btns(3):
'    End Select
'-----------------或者:-----------------
'i = 0
'    If WD.BtnClicked(btns(i)) Then
'            MsgBox "决定d "
'    End If
'------------------------------------
  If WD.BtnClicked <> "btnok" Then MsgBox "决定d "
  
  
'-----------------其他控件------------------
        ctrls = Array("", "", "", "")
        i = 0
            If WD.Res.Exists(ctrls(i)) Then
                MsgBox WD.Res(ctrls(i))
                End If
 End Sub

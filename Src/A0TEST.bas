Attribute VB_Name = "A0TEST"
Private Const mdlname As String = "A0TEST"
Sub Demo_MultiStep_Wizard()

    ' --- Step 1: 选择操作类型 ---
    Dim step1 As Cls_DynaWD
    Set step1 = New Cls_DynaWD
    step1.Title = "Step 1/2 — 选择操作"
    step1.AddUIElement "Label", "lbl_s1", "请选择要执行的操作类型："
    step1.AddUIElement "Button", "btn_copy", "复制子零件", "#43A047"
    step1.AddUIElement "Button", "btn_delete", "删除子零件", "#E53935"
    step1.AddUIElement "Button", "btn_exit", "退出", "#9E9E9E"
    step1.Show vbModal

    Dim action As String
    action = step1.btnClicked
    Set step1 = Nothing

    If action = "btn_exit" Or action = "" Then Exit Sub

    ' --- Step 2: 输入参数 ---
    Dim step2 As Cls_DynaWD
    Set step2 = New Cls_DynaWD
    step2.Title = "Step 2/2 — 配置参数"
    step2.AddUIElement "Label", "lbl_s2", "操作: " & IIf(action = "btn_copy", "复制", "删除")
    step2.AddUIElement "TextBox", "txt_count", "1"
    step2.AddUIElement "CheckBox", "chk_log", "记录操作日志"
    step2.AddUIElement "Button", "btn_run", "执行", "#1E88E5"
    step2.AddUIElement "Button", "btn_back", "返回"
    step2.Show vbModal

    If step2.IsCancelled Or step2.btnClicked = "btn_back" Then
        MsgBox "已返回或取消"
    Else
        Dim cnt As String: cnt = ""
        If step2.Results.Exists("txt_count") Then cnt = step2.Results("txt_count")
        Dim doLog As Boolean: doLog = False
        If step2.Results.Exists("chk_log") Then doLog = CBool(step2.Results("chk_log"))
        MsgBox "执行操作: " & action & vbLf & "数量: " & cnt & vbLf & "记录日志: " & doLog
    End If

    Set step2 = Nothing
End Sub

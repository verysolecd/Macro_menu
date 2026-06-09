Attribute VB_Name = "MDL_translateRename"
Option Explicit
'==============================================================================
' CATIA 结构树名称批量翻译宏
' 翻译引擎: "GOOGLE"(谷歌,默认)
'==============================================================================

Dim gOK  As Long
Dim gErr As Long

'------------------------------------------------------------------------------
' 主程序
'------------------------------------------------------------------------------
Sub 翻译结构树名称()
    gOK = 0: gErr = 0
    ' 收集所有对象（先收集再翻译，避免遍历时改名导致混乱）
    Dim names(5000) As String
    Dim objs(5000)  As Object
    Dim n As Long
    Dim gs
'    Set gs = CATIA.ActiveDocument.part.HybridBodies.item(2)
'    Collect gs, names, objs, n
    For Each gs In CATIA.ActiveDocument.part.HybridBodies
        Collect gs, names, objs, n
    Next
    If n = 0 Then MsgBox "未找到任何几何图形集！", vbExclamation: Exit Sub

    If MsgBox("找到 " & n & " 个名称，引擎：" & "开始翻译？", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Dim i As Long
    For i = 0 To n - 1
        Dim src As String: src = names(i)
        If HasChinese(src) Then
            ' 已含中文，直接跳过（不计入成功/失败）
        Else
            Dim dst As String: dst = Translate(src)
                On Error Resume Next
                objs(i).Name = dst
                If Err.Number = 0 Then gOK = gOK + 1 Else gErr = gErr + 1
                Err.Clear: On Error GoTo 0
        End If
    Next

    On Error Resume Next: CATIA.ActiveDocument.part.Update: On Error GoTo 0
    MsgBox "完成！成功：" & gOK & "  失败：" & gErr, vbInformation
End Sub

'------------------------------------------------------------------------------
' 递归收集几何集及内部图形
'------------------------------------------------------------------------------
Private Sub Collect(gs, names, objs, ByRef n)
    names(n) = gs.Name:
    Set objs(n) = gs:
    n = n + 1

    Dim h
    On Error Resume Next
    For Each h In gs.HybridShapes
        If Err.Number = 0 Then names(n) = h.Name: Set objs(n) = h: n = n + 1
        Err.Clear
    Next
    On Error GoTo 0

    Dim sub_
    For Each sub_ In gs.HybridBodies
        Collect sub_, names, objs, n
    Next
End Sub

'------------------------------------------------------------------------------
' 翻译调度
'------------------------------------------------------------------------------
Private Function Translate(text)
Translate = TransGoogle(text)          ' GOOGLE 及默认
End Function

'------------------------------------------------------------------------------
' Google 翻译（translate.googleapis.com 无需 Key，稳定高频可用）
' 响应格式: [[[ "翻译结果","原文",...],...],...]
'------------------------------------------------------------------------------
Private Function TransGoogle(text)
    On Error GoTo Fail
    Dim url: url = "https://translate.googleapis.com/translate_a/single" & _
                    "?client=gtx&sl=auto&tl=zh-CN&dt=t&q=" & EncodeURL(text)
    Dim xhr: Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", url, False
    xhr.setRequestHeader "User-Agent", "Mozilla/5.0"
    xhr.setRequestHeader "Accept", "application/json"
    xhr.Send
    If xhr.status <> 200 Then GoTo Fail
    Dim res: res = xhr.ResponseText

    ' Google 返回格式: [[["段1","原1",...],["段2","原2",...]],...]
    ' 需拼接所有翻译段（应对长名称被分段的情况）
    Dim combined As String
    Dim pos As Long: pos = 1
    Dim chunks() As String: chunks = Split(res, "[[[""")
    Dim k As Long
    For k = 1 To UBound(chunks)         ' 从 1 开始，跳过头部
        Dim seg As String
        seg = Split(chunks(k), """")(0) ' 取该段译文
        If Len(seg) > 0 Then combined = combined & seg
        ' 如果下一个分块不是翻译段（而是元数据），停止
        If InStr(chunks(k), "]]") > 0 Then Exit For
    Next k
    If Len(combined) > 0 Then
        TransGoogle = Trim(combined)
        Exit Function
    End If
Fail:
    TransGoogle = text
End Function



'------------------------------------------------------------------------------
' UTF-8 URL 编码（用 ADODB.Stream，支持中文/多字节字符）
'------------------------------------------------------------------------------
Private Function EncodeURL(text)
    On Error GoTo Fallback
    Dim st: Set st = CreateObject("ADODB.Stream")
    st.Open: st.Type = 2: st.Charset = "UTF-8"
    st.WriteText text
    st.Position = 0: st.Type = 1: st.Position = 3    ' 跳过 BOM
    Dim B() As Byte: B = st.Read: st.Close
    Dim R As String, i As Long, v As Long   ' v 必须用 Long，Integer 在 b(i)>=128 时会溢出为负
    For i = 0 To UBound(B)
        v = CLng(B(i)) And &HFF             ' 确保取无符号 0~255
        If (v >= 65 And v <= 90) Or (v >= 97 And v <= 122) Or _
           (v >= 48 And v <= 57) Or v = 45 Or v = 95 Or v = 46 Then
            R = R & Chr(v)
        Else
            R = R & "%" & Right("0" & Hex(v), 2)
        End If
    Next
    EncodeURL = R: Exit Function
Fallback:
    R = ""   ' 清空脏数据，防止主路径半途失败后残留内容
    Dim j As Long, c As String
    For j = 1 To Len(text)
        c = Mid(text, j, 1)
        If c Like "[0-9A-Za-z_.-]" Then R = R & c Else R = R & "%" & Right("0" & Hex(Asc(c)), 2)
    Next
    EncodeURL = R
End Function

'------------------------------------------------------------------------------
' 是否含中文（U+4E00~U+9FFF）
'------------------------------------------------------------------------------
Private Function HasChinese(s) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If AscW(Mid(s, i, 1)) >= &H4E00 And AscW(Mid(s, i, 1)) <= &H9FFF Then
            HasChinese = True: Exit Function
        End If
    Next
End Function

'------------------------------------------------------------------------------
' 诊断工具：显示 API 原始返回，用于排查翻译失败原因
'------------------------------------------------------------------------------
Sub 诊断翻译API()
    Dim w: w = InputBox("输入英文测试词：", "API 诊断", "hole")
    If w = "" Then Exit Sub

    Dim url: url = "https://translate.googleapis.com/translate_a/single" & _
                    "?client=gtx&sl=auto&tl=zh-CN&dt=t&q=" & EncodeURL(w)
    Dim xhr: Set xhr = CreateObject("MSXML2.XMLHTTP")
    On Error Resume Next
    xhr.Open "GET", url, False
    xhr.setRequestHeader "User-Agent", "Mozilla/5.0"
    xhr.Send
    Dim res: res = xhr.ResponseText
    If Err.Number <> 0 Then res = "网络错误: " & Err.Description
    On Error GoTo 0

    Dim result: result = "(提取失败)"
    If InStr(res, "[[[""") > 0 Then
        result = Split(Split(res, "[[[""")(1), """")(0)
    End If

    MsgBox "测试词：" & w & vbCrLf & "翻译结果：" & result & vbCrLf & vbCrLf & _
           "--- 原始响应（前400字符）---" & vbCrLf & Left(res, 400), vbInformation, "诊断结果"
End Sub

'------------------------------------------------------------------------------
' 诊断工具2：逐步测试「收集→翻译→改名」全流程，精确定位哪步失败
'------------------------------------------------------------------------------
Sub 诊断改名流程()
    ' Step 1: 取第一个几何图形集
    Dim gs
    On Error Resume Next
    Set gs = CATIA.ActiveDocument.part.HybridBodies.item(1)
    If Err.Number <> 0 Or gs Is Nothing Then
        MsgBox "Step1 失败：找不到任何几何图形集" & vbCrLf & "错误: " & Err.Description, vbCritical: Exit Sub
    End If
    On Error GoTo 0
    Dim objName As String: objName = gs.Name
    MsgBox "Step1 OK：找到对象 [" & objName & "]", vbInformation

    ' Step 2: 翻译
    Dim translated As String
    translated = Translate(objName)

    MsgBox "Step2 OK：翻译结果 [" & translated & "]", vbInformation

    ' Step 3: 改名
    Dim confirm As Integer
    confirm = MsgBox("Step3：将 [" & objName & "] 改名为 [" & translated & "]，确认？", vbYesNo + vbQuestion)
    If confirm <> vbYes Then Exit Sub

    On Error Resume Next
    gs.Name = translated
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    If errNum = 0 Then
        CATIA.ActiveDocument.part.Update
        MsgBox "Step3 OK：改名成功！", vbInformation
    Else
        MsgBox "Step3 失败：改名报错" & vbCrLf & _
               "错误号: " & errNum & vbCrLf & _
               "描述: " & errDesc, vbCritical
    End If
End Sub






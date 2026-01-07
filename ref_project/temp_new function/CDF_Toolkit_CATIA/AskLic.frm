VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AskLic 
   Caption         =   "需要注册码"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "AskLic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AskLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegCancel_Click()
Unload AskLic
End Sub

Private Sub cmdRegOK_Click()
If Len(txtRegNo) <> 8 Then
    MsgBox "请输入正确的注册码"
    Exit Sub
End If

Dim RegForever As String
RegForever = Right(txtRegNo.Text, Len(txtRegNo.Text) - 1)
RegForever = Left(RegForever, Len(RegForever) - 1)

If RegForever = CStr(Int(CPUKey() * 1.1)) Then
    Modifylicfile 2, txtRegNo
    MsgBox "谢谢注册!再次运行即可"
    Unload AskLic
    Exit Sub
End If

If IsNumeric(txtRegNo) = False Then
        MsgBox "请输入正确的注册码"
        Exit Sub
        'ElseIf ((ReadlicLine(0) \ 100000) <> (ReadlicLine(2) \ 10000000)) Or ((ReadlicLine(0) Mod 10) <> (ReadlicLine(2) Mod 10)) Then '20200529,V028版修改
        ElseIf ((ReadlicLine(0) \ 100000) <> (txtRegNo \ 10000000)) Or ((ReadlicLine(0) Mod 10) <> (txtRegNo Mod 10)) Then
            MsgBox "请输入正确的注册码"
            Exit Sub
            ElseIf txtRegNo > 99999999 Or txtRegNo < 10000000 Then
                MsgBox "请输入正确的注册码"
                Exit Sub
End If




Dim RegChk
    RegChk = (txtRegNo Mod 10000000) \ 10 - ReadlicLine(0) + 20000000
    Dim YYYY, MM, DD
    YYYY = RegChk \ 10000
    MM = (RegChk Mod 10000) \ 100
    DD = RegChk Mod 100
    RegChk = YYYY & "-" & MM & "-" & DD
 If IsDate(RegChk) = False Then
    MsgBox "请输入正确的注册码"
    Exit Sub
 ElseIf (DateDiff("d", Now(), CDate(RegChk)) > 0) And (DateDiff("d", Now(), CDate(RegChk)) < 370) Then
    Modifylicfile 2, txtRegNo
    MsgBox "谢谢注册!再次运行即可"
    Unload AskLic
 Else
    MsgBox "请输入正确的注册码"
 End If
End Sub





Private Sub lblReqNo_Click()

End Sub

Private Sub lblWarning_Click()

End Sub

Private Sub txtRegNo_Change()

End Sub

Private Sub UserForm_Initialize()
Dim Picfile
Picfile = "WeChatMatrix.jpg"
Picfile = oCATVBA_Folder("Config").Path & "\" & Picfile
imgWeChat.Picture = LoadPicture(Picfile)
lblReqNo.Caption = "一年授权需求码：" & ReadlicLine(0)
'Randomize
lblReqNo2.Caption = "永久授权需求码：" & "A" & CPUKey() & MacKey() 'Int(89 * Rnd + 10)

End Sub


Sub Createlicfile(ByVal lineNo As Integer, Optional ByVal sTextFileName As String) 'default is "gk.gk"
On Error Resume Next
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim sCATVBA_ConfigFolder
   sCATVBA_ConfigFolder = objFSO.GetParentFolderName(oCATVBA_Folder().Path) & "\Config"
    'MsgBox CATVBA_ConfigFolder
    If Not objFSO.FolderExists(sCATVBA_ConfigFolder) Then
        objFSO.CreateFolder (sCATVBA_ConfigFolder)
        
    End If
    
    Dim objStream, txtName, FILE_NAME
    txtName = "gk.gk"
    If sTextFileName <> "" Then
    txtName = sTextFileName
    End If

    Const TristateFalse = 0
    FILE_NAME = sCATVBA_ConfigFolder & "\" & txtName
    Set objStream = objFSO.CreateTextFile(FILE_NAME, True, TristateFalse)
    
    Dim ReqNo, i, txtLine(99)
    Randomize
    ReqNo = Int((299998 - 100000 + 1) * Rnd + 100000)

        For i = 0 To 99
        txtLine(i) = ReqNo
        Next
     txtLine(lineNo) = (ReqNo \ 100000) * 10000000 + (ReqNo + Int(Format(DateAdd("m", 1, Date), "yymmdd"))) * 10 + (ReqNo Mod 10) '第一位是ReqNo第一位，最后一位是ReqNo最后一位，中间是ReqNo+YYMMDD,MM后延了一个月
        For i = 0 To 99
        objStream.writeline txtLine(i)
        Next
    
    objStream.Close
    Dim dsfolder As String
    dsfolder = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes"
    If Not objFSO.FolderExists(dsfolder) Then
        objFSO.CreateFolder (dsfolder)
    End If
    objFSO.CopyFile FILE_NAME, objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\cat.kl"
    objFSO.CopyFile FILE_NAME, objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\"

Set objFSO = Nothing
Set objStream = Nothing
On Error GoTo 0
End Sub

Sub Modifylicfile(ByVal lineNo As Integer, ByVal NewNo As Double, Optional ByVal sTextFileName As String) 'default is "gk.gk"
On Error Resume Next
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim sCATVBA_ConfigFolder
   sCATVBA_ConfigFolder = objFSO.GetParentFolderName(oCATVBA_Folder().Path) & "\Config"
    Dim dsfolder As String
    dsfolder = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes"
    If Not objFSO.FolderExists(dsfolder) Then
        objFSO.CreateFolder (dsfolder)
    End If
    
    Dim objStream, txtName, FILE_NAME, objS
    txtName = "gk.gk"
    If sTextFileName <> "" Then
    txtName = sTextFileName
    End If

    FILE_NAME = sCATVBA_ConfigFolder & "\" & txtName
    '用户可能已经删除sCATVBA文件夹，故可能取不到FILE_NAME
    '增加FILE_NAME所指文件是否存在的判断，如果不存在，则从SpecialFolder中拷贝
    '修改于20200427
    If Not objFSO.FileExists(FILE_NAME) Then
           If Not objFSO.FolderExists(sCATVBA_ConfigFolder) Then
                objFSO.CreateFolder sCATVBA_ConfigFolder
'                objFSO.GetFolder(sCATVBA_ConfigFolder).Attributes = 2  '20200529取消隐藏操作
           End If
     objFSO.GetFile(objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\cat.kl").Attributes = 0
     objFSO.GetFolder(sCATVBA_ConfigFolder).Attributes = 0
     objFSO.CopyFile objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\cat.kl", sCATVBA_ConfigFolder & "\" & txtName
    End If
    '以上修改于20200427
    Set objS = objFSO.GetFile(FILE_NAME)
    objFSO.GetFolder(sCATVBA_ConfigFolder).Attributes = 0  '20200529取消隐藏
    objFSO.GetFile(FILE_NAME).Attributes = 0 '20200529取消隐藏
    objFSO.GetFile(objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\cat.kl").Attributes = 0  '20200529取消隐藏
    Set objStream = objS.OpenAsTextStream(1, 0)
    Dim i, txtLine(99)
 
        For i = 0 To 99
        txtLine(i) = objStream.ReadLine
        Next
        objStream.Close
    txtLine(lineNo) = NewNo
    
        Set objStream = objS.OpenAsTextStream(2, 0)

        For i = 0 To 99
        objStream.writeline txtLine(i)
        Next
        objStream.Close

    objFSO.CopyFile FILE_NAME, objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\cat.kl"
    objFSO.CopyFile FILE_NAME, objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\"
  Set objFSO = Nothing
  Set objS = Nothing
  Set objStream = Nothing
On Error GoTo 0
End Sub
Function GetDateSerial(ByVal lineNo As Integer, Optional ByVal sTextFileName As String) 'default is "gk.gk"
GetDateSerial = "2019-12-01"
'On Error Resume Next
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim sCATVBA_ConfigFolder
   sCATVBA_ConfigFolder = objFSO.GetParentFolderName(oCATVBA_Folder().Path) & "\Config"
    
    Dim objStream, txtName, FILE_NAME, objS, FILE_NAME0
    txtName = "cat.kl"
    If sTextFileName <> "" Then
    txtName = sTextFileName
    End If
    
    FILE_NAME0 = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\" & txtName
    'FILE_NAME = sCATVBA_ConfigFolder & "\" & txtName
    
    Set objS = objFSO.GetFile(FILE_NAME0)
    Set objStream = objS.OpenAsTextStream(1, 0)
    Dim i, txtLine(99)
 
        For i = 0 To lineNo
        txtLine(i) = objStream.ReadLine
        Next
        objStream.Close
        'MsgBox txtLine(lineNo)
        GetDateSerial = (txtLine(lineNo) Mod 10000000) \ 10 - txtLine(0) + 20000000
    Dim YYYY, MM, DD
    YYYY = GetDateSerial \ 10000
    MM = (GetDateSerial Mod 10000) \ 100
    DD = GetDateSerial Mod 100
    GetDateSerial = YYYY & "-" & MM & "-" & DD

    'MsgBox GetDateSerial
        
  Set objFSO = Nothing
  Set objS = Nothing
  Set objStream = Nothing
'On Error GoTo 0
End Function
Function ReadlicLine(ByVal lineNo As Integer, Optional ByVal sTextFileName As String) 'default is "gk.gk"
ReadlicLine = ""
On Error Resume Next
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim sCATVBA_ConfigFolder
   sCATVBA_ConfigFolder = objFSO.GetParentFolderName(oCATVBA_Folder().Path) & "\Config"
    
    Dim objStream, txtName, FILE_NAME, objS, FILE_NAME0
    txtName = "cat.kl"
    If sTextFileName <> "" Then
    txtName = sTextFileName
    End If
    
    FILE_NAME0 = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\" & txtName
    'FILE_NAME = sCATVBA_ConfigFolder & "\" & txtName
    
    Set objS = objFSO.GetFile(FILE_NAME0)
    Set objStream = objS.OpenAsTextStream(1, 0)
    Dim i, txtLine(99)
 
        For i = 0 To lineNo
        txtLine(i) = objStream.ReadLine
        Next
        objStream.Close
    ReadlicLine = txtLine(lineNo)

    'MsgBox ReadlicLine
        
  Set objFSO = Nothing
  Set objS = Nothing
  Set objStream = Nothing
On Error GoTo 0
End Function
Function IsLicGood(ByVal lineNo As Integer, Optional ByVal sTextFileName As String) 'default is "gk.gk"
IsLicGood = True
'On Error Resume Next
   Dim objFSO
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Dim sCATVBA_ConfigFolder
   sCATVBA_ConfigFolder = objFSO.GetParentFolderName(oCATVBA_Folder().Path) & "\Config"
    
    Dim txtName, FILE_NAME0, FILE_NAME1
    txtName = "cat.kl"
    If sTextFileName <> "" Then
    txtName = sTextFileName
    End If
    Dim dsfolder As String
    dsfolder = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes"
    If Not objFSO.FolderExists(dsfolder) Then
        objFSO.CreateFolder (dsfolder)
    End If
    FILE_NAME0 = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\DassaultSystemes\" & txtName
    FILE_NAME1 = objFSO.GetParentFolderName(objFSO.GetSpecialFolder(2)) & "\gk.gk"
    
    
    
            If objFSO.FileExists(FILE_NAME0) = False Then  '若不存在licfile(首次运行）,则生成licfile
              If objFSO.FileExists(FILE_NAME1) Then
                 objFSO.CopyFile FILE_NAME1, FILE_NAME0
              Else
                Createlicfile 2
                IsLicGood = True
                Exit Function      '否则将licfile读取的日期与当前日期比较，未过期运行，过期则弹出对话框
              End If
            End If
            Dim ReadL As String
            ReadL = ReadlicLine(2)
            ReadL = Right(ReadL, Len(ReadL) - 1)
            ReadL = Left(ReadL, Len(ReadL) - 1)
            If ReadL = CStr(Int(CPUKey() * 1.1)) Then
                 IsLicGood = True
                 Exit Function
            ElseIf IsDate(GetDateSerial(2)) = False Then
                Exit Function
                ElseIf DateDiff("d", Now(), CDate(GetDateSerial(2))) > 0 Then
                    IsLicGood = True
            End If
 
  Set objFSO = Nothing
'On Error GoTo 0
End Function
Function CPUKey()     'Get user CPU serialnumber and calculate a Key
        Dim cpuInfo, moc, mo, i
        cpuInfo = ""
        Set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
        For Each mo In moc
        cpuInfo = mo.ProcessorId
        'MsgBox "CPU SerialNumber is : " & cpuInfo
        Next
        CPUKey = 0
        For i = 1 To Len(cpuInfo)
            CPUKey = CPUKey + 4 * i * CLng(i * (Asc(Mid(cpuInfo, i, 1))))
        Next
        '确保是6位数整数
        Select Case CPUKey
            Case CPUKey < 100000
                CPUKey = CPUKey + 100000
            Case CPUKey > 999999
                Do While CPUKey > 900000
                    CPUKey = CPUKey - 50000
                Loop
        End Select
        'MsgBox CPUKey
        Set moc = Nothing
End Function
Function MacKey()     'Get user Mac address and calculate a Key
        MacKey = 0

        Dim mo, mc, i
        Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
        For Each mo In mc
            If mo.IPEnabled = True Then
            m = m & mo.MacAddress
            Exit For
            End If
        Next
        m = Right(m, 2)
           For i = 1 To Len(m)
                MacKey = MacKey + Asc(Mid(m, i, 1))
           Next
        MacKey = 10 + MacKey Mod 89
End Function

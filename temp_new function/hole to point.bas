Sub CATMain()
    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
    '======= Ҫ��ѡ��body
    Dim imsg, filter(0)
    imsg = "��ѡ��body"
    filter(0) = "Body"
    Dim obdy
    Set obdy = KCL.SelectElement(imsg, filter).Value
    Set targethb = oPart.HybridBodies.Add()
    targethb.Name = "extracted points"
    If Not obdy Is Nothing Then
            Set holeBody = obdy
            For Each Hole In holeBody.Shapes
            If TypeOf Hole Is Hole Then
                Set skt = Hole.Sketch
                Set Pt = HSF.AddNewPointCoord(0, 0, 0)
                Set ref = oPart.CreateReferenceFromObject(skt)
                Pt.PtRef = ref
                Pt.Name = "Pt_" & i
                targethb.AppendHybridShape Pt
                oPart.InWorkObject = Pt
                oPart.Update
                i = i + 1
            End If
        Next
            MsgBox "��ɣ�" & i & "����", vbInformation
    End If

End Sub


' API��������
Public Declare Function FindWindow Lib "user32" alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Integer 

Sub WaitForCATIADisassembleAndContinue()
    ' ��������
    Dim confirmationWindowHandle As Long 
    Dim progressWindowHandle As Long 
    Dim shellObject As Object 
    ' ����Shell�������ڷ��ͼ�������
    Set shellObject = CreateObject("Wscript.Shell") 
    
    ' ��һ�׶Σ��ȴ��ֽ�ȷ�ϴ��ڳ���
    Dim waitCounter As Integer
    waitCounter = 0
    
    ' �ȴ�ȷ�ϴ��ڣ�ͨ����"ȷ��"��"OK"���ڣ�
    Do While waitCounter < 300  ' ���ȴ�30��
        ' ���Բ��ҳ�����CATIAȷ�ϴ��ڱ���
        confirmationWindowHandle = FindWindow(vbNullString, "ȷ��")  ' ���İ�CATIA
        If confirmationWindowHandle = 0 Then
            confirmationWindowHandle = FindWindow(vbNullString, "OK")  ' Ӣ�İ�CATIA
        End If
        If confirmationWindowHandle = 0 Then
            confirmationWindowHandle = FindWindow(vbNullString, "Disassemble")  ' �ֽ�Ի���
        End If
        
        DoEvents() 
        
        If confirmationWindowHandle <> 0 Then 
            ' ȷ�ϴ������ҵ���������ǰ������ȷ�ϼ�
            SetForegroundWindow(confirmationWindowHandle) 
            
            ' �ȴ�һС��ʱ��ȷ��������ȫ����
            Application.Wait Now + TimeValue("00:00:01")
            
            ' ���ͻس���ȷ��
            shellObject.SendKeys "{ENTER}"
            
            ' Ҳ���Գ��Կո����Tab+�س����
            ' shellObject.SendKeys "{SPACE}"
            ' shellObject.SendKeys "{TAB}{ENTER}"
            
            Exit Do
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    If confirmationWindowHandle = 0 Then
        MsgBox "����δ�ҵ��ֽ�ȷ�ϴ��ڣ����ֶ�ȷ�ϻ��鴰�ڱ���"
        Exit Sub
    End If
    
    ' �ڶ��׶Σ��ȴ����������ڳ��ֲ����
    waitCounter = 0
    
    ' �ȴ����������ڳ���
    Do While waitCounter < 600  ' ���ȴ�60��
        ' ���Բ��ҽ��������ڣ�CATIAͨ����"����"��"Progress"���ڣ�
        progressWindowHandle = FindWindow(vbNullString, "����")  ' ���İ�
        If progressWindowHandle = 0 Then
            progressWindowHandle = FindWindow(vbNullString, "Progress")  ' Ӣ�İ�
        End If
        If progressWindowHandle = 0 Then
            progressWindowHandle = FindWindow(vbNullString, "Processing")  ' �����д���
        End If
        
        DoEvents()
        
        If progressWindowHandle <> 0 Then
            ' �������������ҵ����ȴ������
            Exit Do
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    If progressWindowHandle = 0 Then
        ' ����û�н��������ڣ�ֱ�ӵȴ���������
        MsgBox "δ��⵽���������ڣ��ȴ�5������..."
        Application.Wait Now + TimeValue("00:00:05")
    Else
        ' �����׶Σ��ȴ�������������ʧ����ʾ�ֽ���ɣ�
        waitCounter = 0
        
        Do While waitCounter < 1200  ' ���ȴ�2����
            DoEvents()
            
            ' �������������Ƿ񻹴���
            progressWindowHandle = FindWindow(vbNullString, "����")
            If progressWindowHandle = 0 Then
                progressWindowHandle = FindWindow(vbNullString, "Progress")
            End If
            If progressWindowHandle = 0 Then
                progressWindowHandle = FindWindow(vbNullString, "Processing")
            End If
            
            If progressWindowHandle = 0 Then
                ' ��������������ʧ���ֽ����
                Exit Do
            End If
            
            waitCounter = waitCounter + 1
        Loop
        
        If progressWindowHandle <> 0 Then
            MsgBox "���棺���������ڿ���δ�����رգ���������ִ�к�������"
        End If
    End If
    
    ' ���Ľ׶Σ��ֽ���ɣ�ִ�����ĺ�������
    Call ExecuteYourSubsequentCode()
    
    MsgBox "CATIA�ֽ���������ɣ�����������ִ��"
End Sub

' ���ĺ��������������
Sub ExecuteYourSubsequentCode()
    ' ���������ϣ���ڷֽ���ɺ�ִ�еĴ���
    
    ' ʾ��1����������CATIA����
    Dim catia As Object
    On Error Resume Next
    Set catia = GetObject(, "CATIA.Application")
    If Not catia Is Nothing Then
        ' ����CATIA������������
        ' ���磺ѡ������Ԫ�ء�������������
    End If
    
    ' ʾ��2����¼������־
    Debug.Print "�ֽ���������: " & Now()
    
    ' ʾ��3�����½���״̬
    ' YourForm.ProgressBar.Value = 100
    ' YourForm.StatusLabel.Caption = "�ֽ����"
    
    ' ʾ��4���������������
    ' Call YourNextMacro()
    
    ' ���������ʵ�������޸�����Ĵ���
End Sub

' ��ǿ�汾���Զ���ⴰ�ڱ���
Sub WaitForCATIADisassembleEnhanced()
    Dim shellObject As Object
    Set shellObject = CreateObject("Wscript.Shell")
    
    ' ��⵱ǰCATIA�Ĵ��ڱ���ģʽ
    Dim windowTitles As Variant
    windowTitles = DetectCATIAWindowTitles()
    
    ' �ȴ�ȷ�ϴ���
    If WaitForWindowAndSendKey(windowTitles(0), "{ENTER}", 30) Then
        ' �ȴ��������������
        If WaitForWindowCompletion(windowTitles(1), 120) Then
            ' ִ�к�������
            Call ExecuteYourSubsequentCode()
            MsgBox "�ֽ�����ɹ����"
        Else
            MsgBox "���������ڿ����쳣����������ִ�к�������"
            Call ExecuteYourSubsequentCode()
        End If
    Else
        MsgBox "δ���ҵ�ȷ�ϴ��ڣ����ֶ�����"
    End If
End Sub

' ���CATIA���ڱ���
Function DetectCATIAWindowTitles() As Variant
    Dim titles(1) As String
    
    ' ������CATIA���ڱ���
    titles(0) = "ȷ��"  ' ȷ�ϴ��ڱ���
    titles(1) = "����"  ' ���ȴ��ڱ���
    
    ' ��������������Ӹ���Ĵ��ڱ������߼�
    DetectCATIAWindowTitles = titles
End Function

' ͨ�ô��ڵȴ��Ͱ������ͺ���
Function WaitForWindowAndSendKey(windowTitle As String, keyToSend As String, maxWaitSeconds As Integer) As Boolean
    Dim windowHandle As Long
    Dim shellObject As Object
    Dim waitCounter As Integer
    
    Set shellObject = CreateObject("Wscript.Shell")
    waitCounter = 0
    
    Do While waitCounter < maxWaitSeconds * 10  ' ת��Ϊʮ��֮һ��
        windowHandle = FindWindow(vbNullString, windowTitle)
        DoEvents()
        
        If windowHandle <> 0 Then
            SetForegroundWindow(windowHandle)
            Application.Wait Now + TimeValue("00:00:01")
            shellObject.SendKeys keyToSend
            WaitForWindowAndSendKey = True
            Exit Function
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    WaitForWindowAndSendKey = False
End Function

' �ȴ�������ɺ���
Function WaitForWindowCompletion(windowTitle As String, maxWaitSeconds As Integer) As Boolean
    Dim windowHandle As Long
    Dim waitCounter As Integer
    
    waitCounter = 0
    
    ' �ȵȴ����ڳ���
    Do While waitCounter < maxWaitSeconds * 10
        windowHandle = FindWindow(vbNullString, windowTitle)
        DoEvents()
        
        If windowHandle <> 0 Then Exit Do
        waitCounter = waitCounter + 1
    Loop
    
    If windowHandle = 0 Then Return True  ' ����û�н��ȴ���
    
    ' �ȴ�������ʧ
    waitCounter = 0
    Do While waitCounter < maxWaitSeconds * 10
        DoEvents()
        windowHandle = FindWindow(vbNullString, windowTitle)
        
        If windowHandle = 0 Then
            WaitForWindowCompletion = True
            Exit Function
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    WaitForWindowCompletion = False
End Function
Attribute VB_Name = "ASM_sendDir"
'Attribute VB_Name = "m34_sendDir"
'{GP:3}
'{Ep:sendDir}
'{Caption:���ݵ�·��}
'{ControlTipText:send��ǰ����Ʒ��·��}
'{BackColor:}

Sub sendDir()
    CATIA.DisplayFileAlerts = True
    Dim oDoc
    Set oDoc = CATIA.ActiveDocument
    docPath = oDoc.path
    docName = oDoc.Name
    
    initial = docPath & "\" & docName
    
    Dim pn
    
    If KCL.IsType_Of_T(oDoc, "DrawingDocument") Then
        pn = strbflast(oDoc.Name, ".")
    Else
        pn = oDoc.Product.PartNumber
    End If
    
    fname = rmchn(pn)    '�����������ַ��滻Ϊ&
        
    Dim bckFolderName As String
    
    bckFolderName = KCL.strbflast(fname, "_") & "_" & KCL.timestamp("min")
    
    Dim opath
        opath = KCL.ofParentPath(docPath)
        bckFolder = opath & bckFolderName
    
    If Not KCL.isPathchn(bckFolder) Then
        Set Send = CATIA.CreateSendTo()
        Send.KeepDirectory (1)  '1 keepdir �� 0 no keep dir
        Send.SetInitialFile initial
        Send.SetDirectoryFile bckFolder
        Send.Run
        MsgBox "�Ѿ����ݵ�" & bckFolder
        
    Else
        MsgBox bckFolder & vbNewLine & _
        "  " & vbNewLine & _
         "��Ĳ�Ʒ����Ű����Ƿ��ַ����޷����ݣ�����!"
    End If
End Sub


Function rmchn(inputString) As String
    Dim regEx: Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[\u4e00-\u9fa5]"
    regEx.Global = True
    rmchn = regEx.Replace(inputString, " ")
    Set regEx = Nothing
End Function
Sub mdlog()
    Dim oDoc, currPath
    Set oDoc = CATIA.ActiveDocument
    currPath = IIf(oDoc.path = "", "", oDoc.path)
    mdocPath = currPath & ".md"
End Sub

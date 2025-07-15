Attribute VB_Name = "m34_sendDir"
'Attribute VB_Name = "m34_sendDir"
'{GP:3}
'{Ep:sendDir}
'{Caption:���ݵ�·��}
'{ControlTipText:send��ǰ����Ʒ��·��}
'{BackColor:}

Sub sendDir()
    CATIA.DisplayFileAlerts = True
    Dim odoc
    Set odoc = CATIA.ActiveDocument
    Dim rootprd
    Set rootprd = odoc.product
    dpath = odoc.path
    dName = odoc.Name
    initial = dpath & "\" & dName
    Dim bckFolderName As String
    bckFolderName = rootprd.PartNumber & "_" & KCL.timestamp("d")
    
    Dim opath
    opath = KCL.ofParentPath(dpath)
    Set TCF = CATIA.FileSystem.CreateFolder(bckFolderName)
    
    bckFolder = opath & "\" & bckFolderName
    
    If Not KCL.isPathchn(bckFolder) Then
        Set Send = CATIA.CreateSendTo()
        Send.SetInitialFile initial
        Send.SetDirectoryFile bckFolder
        Send.Run
        MsgBox "�Ѿ����ݵ�" & bckFolder
    Else
        MsgBox bckFolder & vbNewLine & _
        "  " & vbNewLine & _
         "·�������Ƿ��ַ����޷����ݣ�����!"
    End If
End Sub
Sub mdlog()
    Dim odoc, currPath
    Set odoc = CATIA.ActiveDocument
    currPath = IIf(odoc.path = "", "", odoc.path)
mdpath = currPath & ".md"
End Sub


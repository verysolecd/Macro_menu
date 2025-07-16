Attribute VB_Name = "m34_sendDir"
'Attribute VB_Name = "m34_sendDir"
'{GP:3}
'{Ep:sendDir}
'{Caption:备份到路径}
'{ControlTipText:send当前根产品到路径}
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
        MsgBox "已经备份到" & bckFolder
    Else
        MsgBox bckFolder & vbNewLine & _
        "  " & vbNewLine & _
         "路径包含非法字符，无法备份，请检查!"
    End If
End Sub
Sub mdlog()
    Dim odoc, currPath
    Set odoc = CATIA.ActiveDocument
    currPath = IIf(odoc.path = "", "", odoc.path)
mdpath = currPath & ".md"
End Sub


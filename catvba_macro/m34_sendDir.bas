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

    dpath = odoc.path
    dName = odoc.Name
    initial = dpath & "\" & dName
 
    Dim pn
    If KCL.IsType_Of_T(odoc, "DrawingDocument") Then
        pn = strbflast(odoc.Name, ".")
    Else
        pn = odoc.product.PartNumber
    End If
    
    fname = rmchn(pn)    '将所有中文字符替换为&
        
    Dim bckFolderName As String
    bckFolderName = KCL.strbflast(fname, "_") & "_" & KCL.timestamp("d")
    
    Dim opath
    opath = KCL.ofParentPath(dpath)
    
    bckFolder = opath & bckFolderName
    If Not KCL.isPathchn(bckFolder) Then
        Set Send = CATIA.CreateSendTo()
        Send.KeepDirectory (1)  '1 keepdir ， 0 no keep dir
        Send.SetInitialFile initial
        Send.SetDirectoryFile bckFolder
        Send.Run
        MsgBox "已经备份到" & bckFolder
    Else
        MsgBox bckFolder & vbNewLine & _
        "  " & vbNewLine & _
         "你的产品零件号包含非法字符，无法备份，请检查!"
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
    Dim odoc, currPath
    Set odoc = CATIA.ActiveDocument
    currPath = IIf(odoc.path = "", "", odoc.path)
mdpath = currPath & ".md"
End Sub

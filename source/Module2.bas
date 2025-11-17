Attribute VB_Name = "Module2"
Sub explorepath()
Dim thisdir
    ipath = "D:\Coding\Myhub\Macro_menu\oTemp"
    
    Dim FSO As Object: Set FSO = GetFso
        If FSO.FileExists(ipath) Then
            thisdir = ipath
        ElseIf FSO.FolderExists(ipath) Then
            Set Fdl = FSO.GetFolder(ipath)

            For Each file In Fdl.Files
                thisdir = file.path
            Exit For
            Next
        End If
End Sub



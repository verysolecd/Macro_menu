VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'文件夹选择器

Option Explicit

Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As LongPtr

Private Type BROWSEINFO
    hOwner As LongPtr
    pidlRoot As LongPtr
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type


'文件夹选择器显示
'return: 文件夹路径
Function show_folder_picker()

    Dim bi As BROWSEINFO
    Dim pidl As LongPtr
    Dim folderPath As String
    
    With bi
        .hOwner = 0
        .lpszTitle = "请选择文件夹"
        .ulFlags = &H1
    End With
    
    pidl = SHBrowseForFolder(bi)
    If pidl <> 0 Then
        folderPath = Space$(260)
        SHGetPathFromIDList pidl, folderPath
        folderPath = Left(folderPath, InStr(folderPath, Chr(0)) - 1)
        show_folder_picker = folderPath
    End If

End Function

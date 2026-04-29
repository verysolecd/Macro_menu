Attribute VB_Name = "Open_Current_Folder"
#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As LongPtr, ByVal y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal wFlags As LongPtr) As LongPtr
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim winHwnd As Long
#End If




Sub CATMain()
IntCATIA
Err.Clear

On Error Resume Next
Dim s As String
s = oActDoc.Path
If Err.Number <> 0 Then
Exit Sub
End If
If CreateObject("Scripting.FileSystemObject").FolderExists(s) = False Then
MsgBox "文件夹不存在,当前活动文档是否并未保存？", vbInformation, "臭豆腐工具箱CATIA版"
Exit Sub
End If
winHwnd = FindWindow(vbNullString, Right(s, Len(s) - InStrRev(s, "\")))
'MsgBox winHwnd
If winHwnd <> 0 Then
ShowWindow winHwnd, 9
SetWindowPos winHwnd, 0, 100, 100, 800, 600, &H40
'ShowWindow winHwnd, 5
Else
    Shell "explorer.exe " & s, vbNormalFocus
End If

End Sub


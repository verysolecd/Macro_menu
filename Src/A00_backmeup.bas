Attribute VB_Name = "A00_backmeup"
Option Explicit
'
Private Const mdlName As String = "A0_backmeup"
Sub bckUp_Modules()
    Dim fm As New VbaModuleManegerView
    On Error Resume Next
        fm.Show
    On Error GoTo 0
End Sub

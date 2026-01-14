Attribute VB_Name = "A0_backmeup"
Option Explicit

'
Sub CATMain()
    Dim fm As New VbaModuleManegerView
    On Error Resume Next
        fm.Show
    On Error GoTo 0
End Sub

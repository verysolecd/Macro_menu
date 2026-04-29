Attribute VB_Name = "Change_Link_2D"
Option Explicit

Sub CATMain()
        
    IntCATIA

    If (TypeName(oActDoc) <> "DrawingDocument") Then
        MsgBox "此命令只能在工程制图模块下运行", vbInformation, "Information"
        Exit Sub
    End If
    
    Change_Link.Show
    Unload GB_Frame

End Sub


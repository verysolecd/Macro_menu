Attribute VB_Name = "A00_backmeup"
Option Explicit
'
Private Const mdlname As String = "A00_backmeup"
Sub bckUp_Modules()
    Dim fm As New VbaModuleManegerView
    On Error Resume Next
        fm.Show
    On Error GoTo 0
End Sub
Public Sub AddmdlName()
    Dim vbComp As Object
    Dim codemod As Object
    Dim i As Long, startline As Long, prockind As Long
    Dim searchStartLine As Long, searchStartCol As Long
    Dim searchEndLine As Long, searchEndCol As Long
    Dim declText As String, constName As String, procName As String
    constName = "mdlname"
    Dim vbprj As Object: Set vbprj = KCL.GetApc().ExecutingProject.VBProject
    Dim colls: Set colls = vbprj.VBComponents
    For Each vbComp In colls
        Set codemod = vbComp.CodeModule
        For i = 1 To codemod.CountOfLines
            procName = codemod.ProcOfLine(i, prockind)
            If procName <> "" Then
                startline = codemod.ProcBodyLine(procName, prockind)
                searchStartLine = 1
                searchStartCol = 1
                searchEndLine = startline
                searchEndCol = -1
                declText = "Private Const " & constName & " As String = """ & vbComp.name & """"
                ' Use variables for Find arguments
                If Not codemod.Find("Const " & constName, searchStartLine, searchStartCol, searchEndLine, searchEndCol) Then
                    codemod.InsertLines startline, declText
                Else
                    codemod.ReplaceLine searchStartLine, declText
                End If
                Exit For
            End If
        Next i
    Next vbComp
        MsgBox "已增加模组名变量 mdlname", vbInformation, "已增加模组名变量 mdlname"
End Sub

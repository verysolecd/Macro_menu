Attribute VB_Name = "A0TEST"
Private Const mdlname As String = "A0TEST"
 Sub TEST()
    Set mdl = KCL.GetApc().ExecutingProject.VBProject.VBE.Activecodepane.CodeModule
 End Sub
Option Explicit

Sub AddModuleNameToAllComponents()
    Dim vbComp, codemod As Object   ' VBIDE.CodeModule
    Dim i, startline, prockind As Long
    Dim declText, constName, procname As String
    constName = "mdlname"
   
    
    Dim vbprj As Object: Set vbprj = KCL.GetApc().ExecutingProject
     
    Set vbProj = Application.VBE.ActiveVBProject
    Dim colls: Set colls = vbprj.VBProject.VBComponents
    For Each vbComp In colls
        Set codemod = vbComp.CodeModule
        For i = 1 To codemod.CountOfLines
            procname = codemod.ProcOfLine(i, prockind)
            If procname <> "" Then
                startline = codemod.ProcBodyLine(procname, prockind)
                If Not codemod.Find("Const " & constName, 1, 1, startline, -1) Then
                      declText = "Private Const " & constName & " As String = """ & vbComp.Name & """"
                    codemod.InsertLines startline, declText
                    Debug.Print "Inserted " & constName & " in " & vbComp.Name
                Else
                    Debug.Print constName & " already exists in " & vbComp.Name
                End If
                Exit For
            End If
        Next i
    Next vbComp
    MsgBox "完成！模块名称已添加到所有组件。", vbInformation, "Done"
End Sub


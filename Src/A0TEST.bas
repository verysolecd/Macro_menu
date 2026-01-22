Attribute VB_Name = "A0TEST"
 Sub TEST()
' Dim DecCnt, DecCode, mdl
'    Dim Apc As Object: Set Apc = KCL.GetApc()
'    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
'
'    On Error Resume Next
'    Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
'    On Error GoTo 0
    
    Set mdl = KCL.GetApc().ExecutingProject.VBProject.VBE.Activecodepane.codemodule
    
    
    
 End Sub


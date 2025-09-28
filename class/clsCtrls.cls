VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCtrls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public ohint As String


Private Sub txt_Gotfocus()

If txt.Text = ohint Then txt.Text = "": txt.ForeColor = &H0

End Sub
Private Sub txt_Lostfocus()

If txt.Text = "" Then txt.Text = ohint: txt.ForeColor = &H808080

End Sub

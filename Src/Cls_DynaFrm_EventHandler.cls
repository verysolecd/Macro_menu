VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_DynaFrm_EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ControlBtn As MSForms.CommandButton
Attribute ControlBtn.VB_VarHelpID = -1
Private mParentFrm As Object ' To hold the reference to the UserForm
Private mControlID As String ' ID of the button

Public Property Set ParentFrm(ByVal frm As Object)
    Set mParentFrm = frm
End Property

Public Property Let ControlID(ByVal value As String)
    mControlID = value
End Property
Private Sub ControlBtn_Click()
    If Not mParentFrm Is Nothing Then
        mParentFrm.Tag = mControlID  ' Set the ID to the Tag property of the form
        mParentFrm.Hide   ' Hide the form to return control to the calling code
    End If
End Sub

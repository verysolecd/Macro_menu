Attribute VB_Name = "A0TEST"
Private Const mdlname As String = "A0TEST"
Sub tet()

Set oprt = CATIA.ActiveDocument.part
Set bd = oprt.bodies.item(3)

Set hi = KCL.SelectElement("nihao ")

MsgBox "laikai "

End Sub

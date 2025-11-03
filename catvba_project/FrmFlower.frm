VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFlower 
   Caption         =   "Flower"
   ClientHeight    =   10260
   ClientLeft      =   10050
   ClientTop       =   375
   ClientWidth     =   9540.001
   OleObjectBlob   =   "FrmFlower.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "FrmFlower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbHide_Click()
    Me.Hide
End Sub

Private Sub CmdDraw_Click()
    pp = 0
    Call iPos
End Sub

Private Sub CMDdraw2_Click()
pp = 0
pp = Val(FrmFlower.qtydrw2.text)
If pp > 5 Then
pp = 5
End If
Debug.Print pp
Call iPos(pp)

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Op1_Click()
ScrR3.Value = 255
ScrG3.Value = 0
ScrB3.Value = 255
ScrR4.Value = 100
ScrG4.Value = 0
ScrB4.Value = 100
End Sub

Private Sub Op10_Click()
ScrR3.Value = 180
ScrG3.Value = 255
ScrB3.Value = 255
ScrR4.Value = 140
ScrG4.Value = 225
ScrB4.Value = 0
End Sub

Private Sub Op2_Click()
ScrR3.Value = 255
ScrG3.Value = 170
ScrB3.Value = 255
ScrR4.Value = 100
ScrG4.Value = 0
ScrB4.Value = 0
End Sub

Private Sub Op3_Click()
ScrR3.Value = 210
ScrG3.Value = 205
ScrB3.Value = 255
ScrR4.Value = 255
ScrG4.Value = 180
ScrB4.Value = 255
End Sub

Private Sub Op4_Click()
ScrR3.Value = 255
ScrG3.Value = 75
ScrB3.Value = 255
ScrR4.Value = 255
ScrG4.Value = 180
ScrB4.Value = 65
End Sub

Private Sub Op5_Click()
ScrR3.Value = 170
ScrG3.Value = 40
ScrB3.Value = 0
ScrR4.Value = 255
ScrG4.Value = 0
ScrB4.Value = 120
End Sub

Private Sub Op6_Click()
ScrR3.Value = 150
ScrG3.Value = 160
ScrB3.Value = 180
ScrR4.Value = 160
ScrG4.Value = 0
ScrB4.Value = 120
End Sub

Private Sub Op7_Click()
ScrR3.Value = 255
ScrG3.Value = 255
ScrB3.Value = 255
ScrR4.Value = 0
ScrG4.Value = 0
ScrB4.Value = 0
End Sub

Private Sub Op8_Click()
ScrR3.Value = 0
ScrG3.Value = 170
ScrB3.Value = 145
ScrR4.Value = 160
ScrG4.Value = 0
ScrB4.Value = 255
End Sub

Private Sub Op9_Click()
ScrR3.Value = 180
ScrG3.Value = 255
ScrB3.Value = 255
ScrR4.Value = 140
ScrG4.Value = 225
ScrB4.Value = 0
End Sub

Private Sub qtydrw2_Change()

End Sub

Private Sub ScrB1_Change()
Lbl1.BackColor = RGB(ScrR1.Value, ScrG1.Value, ScrB1.Value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.Value) & ", G:" & CStr(ScrG1.Value) & ", B:" & CStr(ScrB1.Value)

End Sub

Private Sub ScrB2_Change()
Lbl2.BackColor = RGB(ScrR2.Value, ScrG2.Value, ScrB2.Value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.Value) & ", G:" & CStr(ScrG2.Value) & ", B:" & CStr(ScrB2.Value)

End Sub

Private Sub ScrB3_Change()
Lbl3.BackColor = RGB(ScrR3.Value, ScrG3.Value, ScrB3.Value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.Value) & ", G:" & CStr(ScrG3.Value) & ", B:" & CStr(ScrB3.Value)

End Sub

Private Sub ScrB4_Change()
Lbl4.BackColor = RGB(ScrR4.Value, ScrG4.Value, ScrB4.Value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.Value) & ", G:" & CStr(ScrG4.Value) & ", B:" & CStr(ScrB4.Value)

End Sub

Private Sub ScrG1_Change()
Lbl1.BackColor = RGB(ScrR1.Value, ScrG1.Value, ScrB1.Value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.Value) & ", G:" & CStr(ScrG1.Value) & ", B:" & CStr(ScrB1.Value)

End Sub

Private Sub ScrG2_Change()
Lbl2.BackColor = RGB(ScrR2.Value, ScrG2.Value, ScrB2.Value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.Value) & ", G:" & CStr(ScrG2.Value) & ", B:" & CStr(ScrB2.Value)

End Sub

Private Sub ScrG3_Change()
Lbl3.BackColor = RGB(ScrR3.Value, ScrG3.Value, ScrB3.Value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.Value) & ", G:" & CStr(ScrG3.Value) & ", B:" & CStr(ScrB3.Value)

End Sub

Private Sub ScrG4_Change()
Lbl4.BackColor = RGB(ScrR4.Value, ScrG4.Value, ScrB4.Value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.Value) & ", G:" & CStr(ScrG4.Value) & ", B:" & CStr(ScrB4.Value)

End Sub

Private Sub ScrR1_Change()
Lbl1.BackColor = RGB(ScrR1.Value, ScrG1.Value, ScrB1.Value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.Value) & ", G:" & CStr(ScrG1.Value) & ", B:" & CStr(ScrB1.Value)

End Sub

Private Sub ScrR2_Change()
Lbl2.BackColor = RGB(ScrR2.Value, ScrG2.Value, ScrB2.Value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.Value) & ", G:" & CStr(ScrG2.Value) & ", B:" & CStr(ScrB2.Value)

End Sub

Private Sub ScrR3_Change()
Lbl3.BackColor = RGB(ScrR3.Value, ScrG3.Value, ScrB3.Value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.Value) & ", G:" & CStr(ScrG3.Value) & ", B:" & CStr(ScrB3.Value)

End Sub

Private Sub ScrR4_Change()
Lbl4.BackColor = RGB(ScrR4.Value, ScrG4.Value, ScrB4.Value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.Value) & ", G:" & CStr(ScrG4.Value) & ", B:" & CStr(ScrB4.Value)

End Sub

Private Sub TxtTPlywood_Change()

End Sub

Private Sub TxtX0_Change()

End Sub

Private Sub UserForm_Activate()
    Lbl1.BackColor = RGB(ScrR1.Value, ScrG1.Value, ScrB1.Value)
    Lbl2.BackColor = RGB(ScrR2.Value, ScrG2.Value, ScrB2.Value)
    Lbl3.BackColor = RGB(ScrR3.Value, ScrG3.Value, ScrB3.Value)
    Lbl4.BackColor = RGB(ScrR4.Value, ScrG4.Value, ScrB4.Value)
    Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.Value) & ", G:" & CStr(ScrG4.Value) & ", B:" & CStr(ScrB4.Value)
    Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.Value) & ", G:" & CStr(ScrG3.Value) & ", B:" & CStr(ScrB3.Value)
    Frame2.Caption = "Ovary; R:" & CStr(ScrR2.Value) & ", G:" & CStr(ScrG2.Value) & ", B:" & CStr(ScrB2.Value)
    Frame1.Caption = "Stem; R:" & CStr(ScrR1.Value) & ", G:" & CStr(ScrG1.Value) & ", B:" & CStr(ScrB1.Value)
End Sub

Private Sub UserForm_Click()

End Sub

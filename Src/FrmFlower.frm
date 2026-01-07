VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFlower 
   Caption         =   "Flower"
   ClientHeight    =   10260
   ClientLeft      =   10050
   ClientTop       =   380
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
pp = Val(FrmFlower.qtydrw2.Text)
If pp > 5 Then
pp = 5
End If
Debug.Print pp
Call iPos(pp)

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Op1_Click()
ScrR3.value = 255
ScrG3.value = 0
ScrB3.value = 255
ScrR4.value = 100
ScrG4.value = 0
ScrB4.value = 100
End Sub

Private Sub Op10_Click()
ScrR3.value = 180
ScrG3.value = 255
ScrB3.value = 255
ScrR4.value = 140
ScrG4.value = 225
ScrB4.value = 0
End Sub

Private Sub Op2_Click()
ScrR3.value = 255
ScrG3.value = 170
ScrB3.value = 255
ScrR4.value = 100
ScrG4.value = 0
ScrB4.value = 0
End Sub

Private Sub Op3_Click()
ScrR3.value = 210
ScrG3.value = 205
ScrB3.value = 255
ScrR4.value = 255
ScrG4.value = 180
ScrB4.value = 255
End Sub

Private Sub Op4_Click()
ScrR3.value = 255
ScrG3.value = 75
ScrB3.value = 255
ScrR4.value = 255
ScrG4.value = 180
ScrB4.value = 65
End Sub

Private Sub Op5_Click()
ScrR3.value = 170
ScrG3.value = 40
ScrB3.value = 0
ScrR4.value = 255
ScrG4.value = 0
ScrB4.value = 120
End Sub

Private Sub Op6_Click()
ScrR3.value = 150
ScrG3.value = 160
ScrB3.value = 180
ScrR4.value = 160
ScrG4.value = 0
ScrB4.value = 120
End Sub

Private Sub Op7_Click()
ScrR3.value = 255
ScrG3.value = 255
ScrB3.value = 255
ScrR4.value = 0
ScrG4.value = 0
ScrB4.value = 0
End Sub

Private Sub Op8_Click()
ScrR3.value = 0
ScrG3.value = 170
ScrB3.value = 145
ScrR4.value = 160
ScrG4.value = 0
ScrB4.value = 255
End Sub

Private Sub Op9_Click()
ScrR3.value = 180
ScrG3.value = 255
ScrB3.value = 255
ScrR4.value = 140
ScrG4.value = 225
ScrB4.value = 0
End Sub

Private Sub qtydrw2_Change()

End Sub

Private Sub ScrB1_Change()
Lbl1.BackColor = RGB(ScrR1.value, ScrG1.value, ScrB1.value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.value) & ", G:" & CStr(ScrG1.value) & ", B:" & CStr(ScrB1.value)

End Sub

Private Sub ScrB2_Change()
Lbl2.BackColor = RGB(ScrR2.value, ScrG2.value, ScrB2.value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.value) & ", G:" & CStr(ScrG2.value) & ", B:" & CStr(ScrB2.value)

End Sub

Private Sub ScrB3_Change()
Lbl3.BackColor = RGB(ScrR3.value, ScrG3.value, ScrB3.value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.value) & ", G:" & CStr(ScrG3.value) & ", B:" & CStr(ScrB3.value)

End Sub

Private Sub ScrB4_Change()
Lbl4.BackColor = RGB(ScrR4.value, ScrG4.value, ScrB4.value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.value) & ", G:" & CStr(ScrG4.value) & ", B:" & CStr(ScrB4.value)

End Sub

Private Sub ScrG1_Change()
Lbl1.BackColor = RGB(ScrR1.value, ScrG1.value, ScrB1.value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.value) & ", G:" & CStr(ScrG1.value) & ", B:" & CStr(ScrB1.value)

End Sub

Private Sub ScrG2_Change()
Lbl2.BackColor = RGB(ScrR2.value, ScrG2.value, ScrB2.value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.value) & ", G:" & CStr(ScrG2.value) & ", B:" & CStr(ScrB2.value)

End Sub

Private Sub ScrG3_Change()
Lbl3.BackColor = RGB(ScrR3.value, ScrG3.value, ScrB3.value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.value) & ", G:" & CStr(ScrG3.value) & ", B:" & CStr(ScrB3.value)

End Sub

Private Sub ScrG4_Change()
Lbl4.BackColor = RGB(ScrR4.value, ScrG4.value, ScrB4.value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.value) & ", G:" & CStr(ScrG4.value) & ", B:" & CStr(ScrB4.value)

End Sub

Private Sub ScrR1_Change()
Lbl1.BackColor = RGB(ScrR1.value, ScrG1.value, ScrB1.value)
Frame1.Caption = "Stem; R:" & CStr(ScrR1.value) & ", G:" & CStr(ScrG1.value) & ", B:" & CStr(ScrB1.value)

End Sub

Private Sub ScrR2_Change()
Lbl2.BackColor = RGB(ScrR2.value, ScrG2.value, ScrB2.value)
Frame2.Caption = "Ovary; R:" & CStr(ScrR2.value) & ", G:" & CStr(ScrG2.value) & ", B:" & CStr(ScrB2.value)

End Sub

Private Sub ScrR3_Change()
Lbl3.BackColor = RGB(ScrR3.value, ScrG3.value, ScrB3.value)
Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.value) & ", G:" & CStr(ScrG3.value) & ", B:" & CStr(ScrB3.value)

End Sub

Private Sub ScrR4_Change()
Lbl4.BackColor = RGB(ScrR4.value, ScrG4.value, ScrB4.value)
Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.value) & ", G:" & CStr(ScrG4.value) & ", B:" & CStr(ScrB4.value)

End Sub

Private Sub TxtTPlywood_Change()

End Sub

Private Sub TxtX0_Change()

End Sub

Private Sub UserForm_Activate()
    Lbl1.BackColor = RGB(ScrR1.value, ScrG1.value, ScrB1.value)
    Lbl2.BackColor = RGB(ScrR2.value, ScrG2.value, ScrB2.value)
    Lbl3.BackColor = RGB(ScrR3.value, ScrG3.value, ScrB3.value)
    Lbl4.BackColor = RGB(ScrR4.value, ScrG4.value, ScrB4.value)
    Frame4.Caption = "Inner Petal; R:" & CStr(ScrR4.value) & ", G:" & CStr(ScrG4.value) & ", B:" & CStr(ScrB4.value)
    Frame3.Caption = "Outer Petal; R:" & CStr(ScrR3.value) & ", G:" & CStr(ScrG3.value) & ", B:" & CStr(ScrB3.value)
    Frame2.Caption = "Ovary; R:" & CStr(ScrR2.value) & ", G:" & CStr(ScrG2.value) & ", B:" & CStr(ScrB2.value)
    Frame1.Caption = "Stem; R:" & CStr(ScrR1.value) & ", G:" & CStr(ScrG1.value) & ", B:" & CStr(ScrB1.value)
End Sub

Private Sub UserForm_Click()

End Sub

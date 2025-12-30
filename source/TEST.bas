Attribute VB_Name = "TEST"

Sub idn()
MsgBox canrefresh

End Sub


Private Function canrefresh()


     
     Dim b: b = CATIA.RefreshDisplay

     canrefresh = IIf(b, True, False)
 
    
End Function

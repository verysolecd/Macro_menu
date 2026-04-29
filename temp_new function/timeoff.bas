Public TimeOnOFF As Boolean
Public LastSaveTime As Date

Sub CATMain()
On Error Resume Next

TimeOnOFF = Not TimeOnOFF

If TimeOnOFF Then
    ' 记录开始时间
    LastSaveTime = Now
    
    Dim S As Integer
    While TimeOnOFF = True
        If Second(Now) > S Or Second(Now) = 0 Then
            S = Second(Now)
            
            ' 检查是否已经过了10分钟
            If DateDiff("n", LastSaveTime, Now) >= 10 Then
                ' 保存当前文档
                CATIA.ActiveDocument.Save
                MsgBox CATIA.ActiveDocument.Name & " 已保存 (" & Time & ")"
                
                ' 更新最后保存时间
                LastSaveTime = Now
            End If
        End If
        DoEvents
    Wend
Else
    MsgBox "自动保存已关闭"
End If
End Sub

'----------------------------------------------------------------------------
' Macro:    CatiaV5-DelateDeactivatedElements.catvbs
' Version:  0.0
' Code:     Catia VBS
' Purpose:  
' Autor:    Krzysztof Górzyński
' Datum:    24/03/2015
'----------------------------------------------------------------------------
Sub CATMain()
On Error Resume Next

Dim partDocument As Document
Set partDocument = CATIA.ActiveDocument

Dim myPart As Part
Set myPart = partDocument1.Part

'If Err.Number = 0 Then

    Dim selection1 As Selection
    Set selection1 = partDocument.Selection
    selection1.Search "CATPrtSearch.PartDesign Feature.Activity=FALSE"
    
    If selection1.Count = 0 Then
        MsgBox "Nie ma deaktywowanych elementów"
        Exit Sub
        
    Else
        MsgBox ("Liczba deaktywowanych elementów to: " & selection1.Count & ". Kliknij Tak aby potwierdzi? usuwanie lub Nie aby wyj??.")
        selection1.Delete
        part1.Update
    End If

    
'Else
'MsgBox "Otwary dokument nie jest dokumentem typu PartDesign!"
'End If
End Sub
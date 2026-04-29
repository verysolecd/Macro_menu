Attribute VB_Name = "View_Name_GB_2D"
Option Explicit

Sub CATMain()
    'On Error Resume Next
    IntCATIA
    Dim oDocument
    Set oDocument = CATIA.ActiveDocument
    If TypeName(oDocument) <> "DrawingDocument" Then
        MsgBox "此命令的功能是将工程图的剖视图，断面视图，细节视图，向视图，轴测视图按中国国标命名，因此需要在工程制图环境下运行！" & vbCrLf & _
                     "运行此宏命令后，在工程制图的激活视图中单击鼠标左键以将视图名称放置在指定位置即可！" & vbCrLf & vbCrLf & _
                     "请在工程制图环境下运行此命令，现在退出...", vbInformation, "臭豆腐工具箱CATIA版”"
        Exit Sub
        
    End If
    Dim oDrawingDoc2 As Object 'DrawingDocument
    Set oDrawingDoc2 = oDocument
    
    Dim oDrwSheet As Object
    Set oDrwSheet = oDocument.Sheets.ActiveSheet
    
    Dim oDrwViews As Object 'DrawingViews
    Set oDrwViews = oDrwSheet.Views
    
    Dim oDrwView As Object 'DrawingView
    Set oDrwView = oDrwViews.ActiveView
    
    Dim oDrawingTexts As Object 'DrawingTexts
    Set oDrawingTexts = oDrwView.Texts
    
    Dim oDrawingText As Object 'DrawingText
    
    Dim iParameters As Object 'Parameters
    Set iParameters = oDrawingDoc2.Parameters
    Dim iParameter As Object 'Parameter
    
    Dim MyPrefix As String, MyIdent As String, MySuffix As String
    Dim oPrefix As String, oIdent As String, oSuffix As String
    Dim DrawingWindowLocation(1)
    Dim Status
    Dim itext As String

    Dim ViewNo As Integer
    Dim FrontViewScale
    
    If oDrwSheet.Views.Count <> 0 Then
        For ViewNo = 1 To oDrwSheet.Views.Count
            If oDrwSheet.Views.Item(ViewNo).ViewType = 1 Then 'catViewFront
                FrontViewScale = oDrwSheet.Views.Item(ViewNo).Scale2
                GoTo Skip
            Else
                FrontViewScale = 0
            End If
        Next ViewNo
    End If
    
Skip:
    If oDrwView.ViewType = 9 Or oDrwView.ViewType = 10 Then      '>>>>剖视图和断面视图<<<<<
        Status = oDocument.Indicate2D("在当前激活视图中单击鼠标左键以将视图名称放置在指定位置", DrawingWindowLocation)
        If Status = "Cancel" Then Exit Sub
                    
        oDrwView.GetViewName MyPrefix, MyIdent, MySuffix
    
        oPrefix = ""
        oSuffix = ""
        oDrwView.SetViewName oPrefix, MyIdent, oSuffix
        
        If oDrwView.Scale2 <> FrontViewScale Then
        itext = vbLf
        Else: itext = ""
        End If
        
        Set oDrawingText = oDrawingTexts.Add(itext, DrawingWindowLocation(0), DrawingWindowLocation(1))
        
        Set iParameter = iParameters.Item(oDrwSheet.Name & "\" & oDrwView.Name & "\Name")
        oDrawingText.InsertVariable 1, 0, iParameter
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrwView.InsertViewScale Len(oDrawingText.Text) + 1, oDrawingText
        End If
        
        oDrawingText.SetParameterOnSubString 0, 0, 0, 1                           '设为粗体catBold
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrawingText.SetParameterOnSubString 2, 0, Len(oDrwView.Name) + 1, 1 '名称部分加下划线
        End If
        
        oDrawingText.SetParameterOnSubString 13, 0, 0, 2                      '设为右侧对齐2,居中1,靠左0
        oDrawingText.SetFontSize 0, 0, 3.5                                              '字体改为3.5
        
    ElseIf oDrwView.ViewType = 11 Then        '>>>>细节视图<<<<<catViewDetail
        
        Dim i As Integer
        Static sequence As Integer
        sequence = 0
        Dim SequenceOfView As Integer
        
        If oDrwSheet.Views.Count <> 0 Then
            For i = 1 To oDrwSheet.Views.Count
            If oDrwSheet.Views.Item(i).ViewType = 11 Then
            sequence = sequence + 1
                If oDrwSheet.Views.Item(i).Name = oDrwView.Name Then
                SequenceOfView = sequence
                End If
            End If
            Next i
        End If
        
        Status = oDocument.Indicate2D("在当前激活视图中单击鼠标左键以将视图名称放置在指定位置", DrawingWindowLocation)
        If Status = "Cancel" Then Exit Sub
        
        Select Case SequenceOfView
        Case 1
            oIdent = "Ⅰ"
        Case 2
            oIdent = "Ⅱ"
        Case 3
            oIdent = "Ⅲ"
        Case 4
            oIdent = "Ⅳ"
        Case 5
            oIdent = "Ⅴ"
        Case 6
            oIdent = "Ⅵ"
        Case 7
            oIdent = "Ⅶ"
        Case 8
            oIdent = "Ⅷ"
        Case 9
            oIdent = "Ⅸ"
        Case 10
            oIdent = "Ⅹ"
        Case 11
            oIdent = "Ⅺ"
        Case 12
            oIdent = "Ⅻ"
        End Select
        
        oPrefix = " "
        oSuffix = " "
        oDrwView.SetViewName oPrefix, oIdent, oSuffix
        
        If oDrwView.Scale2 <> FrontViewScale Then
        itext = vbLf
        Else: itext = ""
        End If
                
        Set oDrawingText = oDrawingTexts.Add(itext, DrawingWindowLocation(0), DrawingWindowLocation(1))
        
        Set iParameter = iParameters.Item(oDrwSheet.Name & "\" & oDrwView.Name & "\Name")
        oDrawingText.InsertVariable 1, 0, iParameter
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrwView.InsertViewScale Len(oDrawingText.Text) + 1, oDrawingText
        End If
    
        oDrawingText.SetParameterOnSubString 0, 0, 0, 1
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrawingText.SetParameterOnSubString 2, 0, Len(oDrwView.Name) + 1, 1
        End If
        
        oDrawingText.SetParameterOnSubString 13, 0, 0, 2
        oDrawingText.SetFontSize 0, 0, 3.5

    ElseIf oDrwView.ViewType = 8 Then     '>>>>轴测视图<<<<<
        Status = oDocument.Indicate2D("在当前激活视图中单击鼠标左键以将视图名称放置在指定位置", DrawingWindowLocation)

        If Status = "Cancel" Then Exit Sub
        
        oPrefix = "轴测视图 / ISOMETRIC VIEW"
        oIdent = ""
        oSuffix = ""
        oDrwView.SetViewName oPrefix, oIdent, oSuffix
        
        If oDrwView.Scale2 <> FrontViewScale Then
        itext = vbLf
        Else: itext = ""
        End If
        
        Set oDrawingText = oDrawingTexts.Add(itext, DrawingWindowLocation(0), DrawingWindowLocation(1))
        
        Set iParameter = iParameters.Item(oDrwSheet.Name & "\" & oDrwView.Name & "\Name")
        oDrawingText.InsertVariable 1, 0, iParameter
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrwView.InsertViewScale Len(oDrawingText.Text) + 1, oDrawingText
        End If
    
        oDrawingText.SetParameterOnSubString 0, 0, 0, 1
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrawingText.SetParameterOnSubString 2, 0, Len(oDrwView.Name) + 1, 1
        End If
        
        oDrawingText.SetParameterOnSubString 13, 0, 0, 1
        oDrawingText.SetFontSize 0, 0, 3.5
        
    ElseIf oDrwView.ViewType = 7 Then    '>>>>向视图<<<<<
        Status = oDocument.Indicate2D("在当前激活视图中单击鼠标左键以将视图名称放置在指定位置", DrawingWindowLocation)
        If Status = "Cancel" Then Exit Sub
        
        oDrwView.GetViewName MyPrefix, MyIdent, MySuffix

        oPrefix = "  "
        oIdent = ""
        oSuffix = "  "
        oDrwView.SetViewName oPrefix, MyIdent, oSuffix
        
        If oDrwView.Scale2 <> FrontViewScale Then
        itext = vbLf
        Else: itext = ""
        End If
        
        Set oDrawingText = oDrawingTexts.Add(itext, DrawingWindowLocation(0), DrawingWindowLocation(1))
        
        Set iParameter = iParameters.Item(oDrwSheet.Name & "\" & oDrwView.Name & "\Name")
        oDrawingText.InsertVariable 1, 0, iParameter
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrwView.InsertViewScale Len(oDrawingText.Text) + 1, oDrawingText
        End If
    
        oDrawingText.SetParameterOnSubString 0, 0, 0, 1
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrawingText.SetParameterOnSubString 2, 0, Len(oDrwView.Name) + 1, 1
        End If
        
        oDrawingText.SetParameterOnSubString 13, 0, 0, 2
        oDrawingText.SetFontSize 0, 0, 3.5

    ElseIf oDrwView.ViewType = 15 Then  '>>>>展开视图<<<<<
        Status = oDocument.Indicate2D("在当前激活视图中单击鼠标左键以将视图名称放置在指定位置", DrawingWindowLocation)

        If Status = "Cancel" Then Exit Sub
        
        oPrefix = "展开视图 / UNFOLDED VIEW"
        oIdent = ""
        oSuffix = ""
        oDrwView.SetViewName oPrefix, oIdent, oSuffix
        
        If oDrwView.Scale2 <> FrontViewScale Then
        itext = vbLf
        Else: itext = ""
        End If
        
        Set oDrawingText = oDrawingTexts.Add(itext, DrawingWindowLocation(0), DrawingWindowLocation(1))
        
        Set iParameter = iParameters.Item(oDrwSheet.Name & "\" & oDrwView.Name & "\Name")
        oDrawingText.InsertVariable 1, 0, iParameter
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrwView.InsertViewScale Len(oDrawingText.Text) + 1, oDrawingText
        End If
    
        oDrawingText.SetParameterOnSubString 0, 0, 0, 1
        
        If oDrwView.Scale2 <> FrontViewScale Then
        oDrawingText.SetParameterOnSubString 2, 0, Len(oDrwView.Name) + 1, 1
        End If
        
        oDrawingText.SetParameterOnSubString 13, 0, 0, 1
        oDrawingText.SetFontSize 0, 0, 3.5

    
    End If
        
    sequence = 0
    
End Sub



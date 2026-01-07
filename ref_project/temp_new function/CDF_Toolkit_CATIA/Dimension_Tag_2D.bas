Attribute VB_Name = "Dimension_Tag_2D"
Option Explicit
'Dim txtStartno As Integer
'Dim txtNextno As Integer
'Dim chkRefit As Boolean
Dim sFontstyle As CatTextProperty
Const PreFix = "MyDim_"
Public sFinNum
Dim DrawingSheets1 As DrawingSheets
Dim drawingSheet1 As DrawingSheet
Dim drawingViews1 As DrawingViews
Dim drawingView1 As DrawingView
Dim DrwText

'************************************设置快捷键***********************************
'************************************设置快捷键***********************************
#If VBA7 Then
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr

Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function RegisterHotKey Lib "user32" (ByVal hwnd As LongPtr, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Private Declare PtrSafe Function UnregisterHotKey Lib "user32" (ByVal hwnd As LongPtr, ByVal id As Long) As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Dim key_preWinProc As LongPtr '用来保存窗口信息
#Else
'Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim key_preWinProc As Long '用来保存窗口信息


#End If

Const WM_HOTKEY = &H312
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const GWL_WNDPROC = (-4)    '窗口函数的地址

'Dim key_preWinProc As LongPtr '用来保存窗口信息
Dim key_Modifiers As Long, key_uVirtKey As Long, key_idHotKey As Long
Dim key_IsWinAddress    As Boolean '是否取得窗口信息的判断
Sub CATMain()
        
IntCATIA
If (TypeName(oActDoc) <> "DrawingDocument") Then
    MsgBox "此命令只能在工程制图模块下运行", vbInformation, "Information"
    Exit Sub
End If
Dim_Tag.Show vbModeless
Dim_Tag.Left = Dim_Tag.Left * 2

End Sub

#If VBA7 Then
Function keyWndproc(ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Function keyWndproc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
      If msg = WM_HOTKEY Then
          Select Case wParam 'wParam 值就是 key_idHotKey
              Case 1 '激活 3 个热键后,3 个热键所对应的操作,大家在其他的程序中，只要修改此处就可以了
                        ManualNo Dim_Tag.txtNextno.Value
                        Dim_Tag.txtNextno.Value = Dim_Tag.txtNextno.Value + 1
              Case 2
                   'MsgBox "F5"
                   RefreshInfo
                'Dim_Tag.cmdRefresh_Click
             Case 3
                 'MsgBox "NC程序改名输出快捷键"
                 MProg.cmdSel_Click
          End Select
      End If
    
      '将消息传送给指定的窗口
      keyWndproc = CallWindowProc(key_preWinProc, hwnd, msg, wParam, lParam)
    
End Function

Function SetHotkey(ByVal KeyId As Long, ByVal KeyAss0 As String, ByVal Action As String, ByVal WinName As String)
'     Dim WinName As String
'         WinName = "UserForm5"
      Dim KeyAss1 As Long
      Dim KeyAss2 As String
      Dim i As Long
    
      i = InStr(1, KeyAss0, ",")
      If i = 0 Then
          KeyAss1 = Val(KeyAss0)
          KeyAss2 = ""
      Else
          KeyAss1 = Right(KeyAss0, Len(KeyAss0) - i)
          KeyAss2 = Left(KeyAss0, i - 1)
      End If
    
      key_idHotKey = 0
      key_Modifiers = 0
      key_uVirtKey = 0
    
      If key_IsWinAddress = False Then    '判断是否需要取得窗口信息，如果重复取得,再最后恢复窗口时，将会造成程序死掉
          '记录原来的window程序地址
          'MsgBox FindWindow(vbNullString, WinName)
          'MsgBox GWL_WNDPROC
          ' This VBA Macro Developed by Charles.Tang
          ' WeChat Chtang80,CopyRight reserved
          key_preWinProc = GetWindowLongPtr(FindWindow(vbNullString, WinName), GWL_WNDPROC)
          SetWindowLongPtr FindWindow(vbNullString, WinName), GWL_WNDPROC, AddressOf keyWndproc

      End If

      key_idHotKey = KeyId
      Select Case Action
          Case "Add"
              If KeyAss2 = "Ctrl" Then key_Modifiers = MOD_CONTROL
              If KeyAss2 = "Alt" Then key_Modifiers = MOD_ALT
              If KeyAss2 = "Shift" Then key_Modifiers = MOD_SHIFT
              If KeyAss2 = "Ctrl+Alt" Then key_Modifiers = MOD_CONTROL + MOD_ALT
              If KeyAss2 = "Ctrl+Shift" Then key_Modifiers = MOD_CONTROL + MOD_SHIFT
              If KeyAss2 = "Ctrl+Alt+Shift" Then key_Modifiers = MOD_CONTROL + MOD_ALT + MOD_SHIFT
              If KeyAss2 = "Shift+Alt" Then key_Modifiers = MOD_SHIFT + MOD_ALT
              key_uVirtKey = Val(KeyAss1)
              RegisterHotKey FindWindow(vbNullString, WinName), key_idHotKey, key_Modifiers, key_uVirtKey '向窗口注册系统热键
              key_IsWinAddress = True '不需要再取得窗口信息
            
          Case "Del"
              SetWindowLongPtr FindWindow(vbNullString, WinName), GWL_WNDPROC, key_preWinProc '恢复窗口信息
              UnregisterHotKey FindWindow(vbNullString, WinName), key_uVirtKey '取消系统热键
              key_IsWinAddress = False '可以再次取得窗口信息
      End Select
End Function
'************************************设置快捷键***********************************
'************************************设置快捷键***********************************


Sub AutoNum(startno As Integer)

Dim Mytexts, MyText, MyNum, MyDims, MyDim, vw, dm, dmbox(7), i
Dim DicY, DicV
Set DicV = CreateObject("Scripting.Dictionary") 'Dictionary for drawingviews
Set DicY = CreateObject("Scripting.Dictionary") 'Dictionary for BalloonNumbers

MyNum = startno - 1
sFinNum = MyNum

Set DrawingSheets1 = oActDoc.Sheets
 
    Set drawingSheet1 = DrawingSheets1.ActiveSheet
    Set drawingViews1 = drawingSheet1.Views
        For vw = 1 To drawingViews1.Count

            Set drawingView1 = drawingViews1.Item(vw)
            Set MyDims = drawingView1.Dimensions
            Set Mytexts = drawingView1.Texts

                For dm = 1 To MyDims.Count

                    MyNum = MyNum + 1
                    'MsgBox MyDims.Count
                    Set MyDim = MyDims.Item(dm)
                    MyDim.GetBoundaryBox dmbox
                    If Abs(dmbox(2) - dmbox(4)) > Abs(dmbox(3) - dmbox(5)) Then '横放矩形
                    Set MyText = Mytexts.Add(CStr(MyNum), dmbox(6) + 2 / drawingView1.Scale2, dmbox(7) - 1 / drawingView1.Scale2)
                    MyText.AnchorPosition = 2 'catMiddleLeft
                    Else   '竖放矩形，可能是竖放尺寸
                    Set MyText = Mytexts.Add(CStr(MyNum), dmbox(4) - 2 / drawingView1.Scale2, dmbox(3) + 1 / drawingView1.Scale2)
                    MyText.AnchorPosition = 6 'catBottomCenter
                    End If
                    MyText.Name = PreFix & "a" & CStr(MyNum)
                    MyText.AssociativeElement = MyDim
                    MyText.Angle = MyDim.ValueAngle  '文字方向
                    FormatDimTag MyText
                    'DicY.Add CStr(MyNum), Mytext
                    'MsgBox Mytext.X & "||" & Mytext.Y

                Next
            Set MyDims = drawingView1.Weldings
                For dm = 1 To MyDims.Count
                    MyNum = MyNum + 1
                    Set MyText = Mytexts.Add(CStr(MyNum), MyDims.Item(dm).x - 5 / drawingView1.Scale2, MyDims.Item(dm).y)
                    MyText.Name = PreFix & "a" & CStr(MyNum)
                    MyText.AssociativeElement = MyDims.Item(dm)
                    FormatDimTag MyText
                Next
            Set MyDims = Nothing
            Set Mytexts = Nothing
            Set drawingView1 = Nothing

         Next

For i = 1 To drawingViews1.Count
      DicV.Add CStr(i), drawingViews1.Item(i)
Next
                                    
     
While DicV.Count <> 0
  Set drawingView1 = oMinDisV(0, 1000, DicV) '与左上角距离最近的视图
  Dim DrawingTexts, j
  Set DrawingTexts = drawingView1.Texts
      For j = 1 To DrawingTexts.Count
            If Left(DrawingTexts.Item(j).Name, Len(PreFix)) = PreFix Then
            DicY.Add j, DrawingTexts.Item(j)
            End If
      Next
      
  SortT 0, 1000, DicY
Wend
        
MsgBox MyNum - startno + 1 & " 个尺寸编号已添加" 'Dimension Identification added!"

End Sub

Private Sub Sort(a() As Integer)
Dim i As Integer, j As Integer, T As Integer
For i = LBound(a()) To UBound(a()) - 1
    For j = LBound(a()) + 1 To UBound(a())
        If a(j - 1) > a(j) Then
        T = a(j - 1)
        a(j - 1) = a(j)
        a(j) = T
        End If
    Next
Next
End Sub

Sub SortAll()
Dim sh, i, DrawingSheets1, drawingSheet1, drawingViews1, drawingView1
Dim DicY, DicV
Set DicV = CreateObject("Scripting.Dictionary") 'Dictionary for drawingviews
Set DicY = CreateObject("Scripting.Dictionary") 'Dictionary for BalloonNumbers

Set DrawingSheets1 = oActDoc.Sheets
For sh = 1 To DrawingSheets1.Count

sFinNum = 0
    Set drawingSheet1 = DrawingSheets1.Item(sh)
    Set drawingViews1 = drawingSheet1.Views
    
        For i = 1 To drawingViews1.Count
              DicV.Add CStr(i), drawingViews1.Item(i)
        Next
        
        While DicV.Count <> 0
          Set drawingView1 = oMinDisV(0, 1000, DicV) '与左上角距离最近的视图
          Dim DrawingTexts, j
          Set DrawingTexts = drawingView1.Texts
              For j = 1 To DrawingTexts.Count
                    If Left(DrawingTexts.Item(j).Name, Len(PreFix)) = PreFix Then
                    DicY.Add j, DrawingTexts.Item(j)
                    End If
              Next
              
          SortT 0, 1000, DicY
        Wend
Next
End Sub
Sub RefreshInfo()
IntCATIA
 Dim_Tag.lstNumbers.Clear

         Dim colDimsA As New VBA.Collection
         Dim colDimsLost As New VBA.Collection
         Dim colDimsDuplicated As New VBA.Collection
         Dim n, j
         Set colDimsA = colSearchName("MyDim_")
        
        j = 1

        n = colDimsA.Count
Dim_Tag.lblTagno.Caption = "编号个数: " & n
        
        While n <> 0

            Dim k, i
            k = 0 '重复编号计数
            
             For Each i In colDimsA
             'MsgBox "typename i is " & TypeName(i)
             'MsgBox "j=" & j & vbCrLf & "colDimsA.Count=" & colDimsA.Count & vbCrLf & "i=" & i & vbCrLf & "Int(colDimsA.Item(i).Text) is " & colDimsA.Item(i).Text 'Int(colDimsA.Item(i).Text)
             Select Case k
             Case 0
                If j = Int(i.Text) Then
                Dim_Tag.lstNumbers.AddItem j
                k = k + 1

                End If
            Case Else
                If j = Int(i.Text) Then
                Dim_Tag.lstNumbers.AddItem j & " -重复" & k & "次"
                k = k + 1
                colDimsDuplicated.Add i
                End If
             End Select

             Next
             
            If k = 0 Then
                Dim_Tag.lstNumbers.AddItem j & " -缺失"
                colDimsLost.Add j
              
            End If
             
             j = j + 1
             n = n - k

        Wend
        
Dim_Tag.lblMaxno.Caption = "最大编号: " & j - 1
Dim_Tag.txtNextno.Value = j
'Dim_Tag.Show vbModeless
If colDimsDuplicated.Count <> 0 Then
MsgBox "发现有 " & colDimsDuplicated.Count & " 处编号重复!"
End If
'如果有编号注释则更新编号注释
Dim sLost As String
sLost = ""
If colDimsLost.Count <> 0 Then
Dim R As Integer
For R = 1 To colDimsLost.Count
sLost = sLost & colDimsLost.Item(R) & "/"
Next
End If

        Dim ExistingNums1 As New VBA.Collection
        Set ExistingNums1 = colSearchName("MyNumNote")
        If ExistingNums1.Count <> 0 Then
                      Dim i1
                         For Each i1 In ExistingNums1
                            i1.Text = "#代表尺寸编号, 共有编号 " & colDimsA.Count & "个, " & "最大的编号是 " & j - 1 & vbCrLf & Space(7) & "缺失的编号是: " & sLost
                            i1.SetParameterOnSubString catBorder, 0, 0, catNone
                            i1.SetParameterOnSubString catBorder, 1, 1, catOblong
                         Next

        End If


End Sub
Function oMinDisV(x0, y0, DicX)

  
Dim i, arrDicX_keys, sMinDisKey, DisTemp, DisCal, AbsCo1(1)

DisTemp = 9999


arrDicX_keys = DicX.keys
sMinDisKey = arrDicX_keys(0)


'找出最小距离对应的字典键值
For i = 0 To UBound(arrDicX_keys)
       AbsCo1(0) = DicX.Item(arrDicX_keys(i)).x
       AbsCo1(1) = DicX.Item(arrDicX_keys(i)).y
       DisCal = Distance2P(x0, y0, AbsCo1)
         If DisTemp > DisCal Then
                DisTemp = DisCal
                sMinDisKey = arrDicX_keys(i)
         End If
Next

Set oMinDisV = DicX.Item(sMinDisKey)
DicX.Remove sMinDisKey
End Function
Sub SortT(x0, y0, oDicT) '连续取最小距离点，迭代
If oDicT.Count = 0 Then
Exit Sub
End If

Dim i, arrDicT_keys, oMinDis, sMinDisKey, DisTemp, AbsCo1(1), RelCo1(1), j

DisTemp = 9999


arrDicT_keys = oDicT.keys
sMinDisKey = arrDicT_keys(0)


'MsgBox "Dict has " & UBound(arrDicT_keys)

'找出最小距离对应的字典键值
For i = 0 To UBound(arrDicT_keys)

        RelCo1(0) = oDicT.Item(arrDicT_keys(i)).x
        RelCo1(1) = oDicT.Item(arrDicT_keys(i)).y
        
    CatAbsoluteCoordinates oDicT.Item(arrDicT_keys(i)).Parent.Parent, AbsCo1, RelCo1

    If DisTemp > Distance2P(x0, y0, AbsCo1) Then
        DisTemp = Distance2P(x0, y0, AbsCo1)
        sMinDisKey = arrDicT_keys(i)
        
    End If
 
Next
Set oMinDis = oDicT.Item(sMinDisKey)

        RelCo1(0) = oMinDis.x
        RelCo1(1) = oMinDis.y
        
    CatAbsoluteCoordinates oMinDis.Parent.Parent, AbsCo1, RelCo1

'MsgBox "Min distance is " & DisTemp & vbCrLf & "X=" & AbsCo1(0) & vbCrLf & "Y=" & AbsCo1(1)
sFinNum = sFinNum + 1
oMinDis.Text = CStr(sFinNum)
oMinDis.Name = PreFix & "s" & CStr(sFinNum)

oDicT.Remove sMinDisKey
'MsgBox oDicT.Count

If oDicT.Count = 0 Then
Exit Sub
Else
Call SortT(AbsCo1(0), AbsCo1(1), oDicT)
End If
End Sub
Sub FormatDimTag(ByVal DimTag As DrawingText)
'DimTag.SetFontName 0, 0, "Arial Narrow"  '"Courrier 10 BT"
DimTag.SetFontSize 0, 0, 2#
DimTag.SetParameterOnSubString catBorder, 0, 0, catOblong
'DimTag.TextProperties.FrameType = catEllipse 'catOblong  'catEllipse
DimTag.TextProperties.Update
End Sub
Function Distance2P(x11, y11, AbsCo22) '两点距离
Distance2P = Sqr((AbsCo22(0) - x11) * (AbsCo22(0) - x11) + (AbsCo22(1) - y11) * (AbsCo22(1) - y11))
End Function
Private Sub CatAbsoluteCoordinates(CatDrawingView As Object, AbsoluteCoordinates(), RelativeCoordinates())

' Compute the coordinates of a point in a view according to the sheet's reference axis
' Location, Angle and Scale factor of the view are take into account
AbsoluteCoordinates(0) = CatDrawingView.xAxisData + (RelativeCoordinates(0) * Cos(CatDrawingView.Angle) - RelativeCoordinates(1) * Sin(CatDrawingView.Angle)) * CatDrawingView.Scale2
AbsoluteCoordinates(1) = CatDrawingView.yAxisData + (RelativeCoordinates(0) * Sin(CatDrawingView.Angle) + RelativeCoordinates(1) * Cos(CatDrawingView.Angle)) * CatDrawingView.Scale2
AbsoluteCoordinates(0) = Round(AbsoluteCoordinates(0), 0)
AbsoluteCoordinates(1) = Round(AbsoluteCoordinates(1), 0)
End Sub

Sub ManualNo(nextno)
Set oSel = oActDoc.Selection
oSel.Clear

Dim DrawingView
Dim DrawingTexts
Dim ManuNo
 Set DrawingView = oActDoc.Sheets.ActiveSheet.Views.ActiveView
 Set DrawingTexts = DrawingView.Texts
 'We propose to the user that he specify a location in the drawing window
 Dim DrawingWindowLocation(1)
 Dim Status
 Status = oActDoc.Indicate2D("鼠标单击放置手动编号,下一编号是 " & nextno & ", Esc 退出", DrawingWindowLocation)
 If (Status = "Cancel") Then Exit Sub
 Set ManuNo = DrawingTexts.Add(nextno, DrawingWindowLocation(0), DrawingWindowLocation(1))
 ManuNo.Name = PreFix & "s" & nextno
 FormatDimTag ManuNo
 
'txtNextno.Value = txtNextno.Value + 1
Set DrawingView = Nothing
Set DrawingTexts = Nothing
End Sub
Function colSearchName(ByVal sName As String)

Dim colSearchNameA As New VBA.Collection
Dim drawingSheet0, drawingViews0, drawingView0, Mytexts0, ts

Set drawingSheet0 = oActDoc.Sheets.ActiveSheet
    Set drawingViews0 = drawingSheet0.Views
        For Each drawingView0 In drawingViews0
            Set Mytexts0 = drawingView0.Texts
                For ts = 1 To Mytexts0.Count
                If Left(Mytexts0.Item(ts).Name, Len(sName)) = sName Then
                colSearchNameA.Add Mytexts0.Item(ts)
                End If
                Next
        Next

Set colSearchName = colSearchNameA
End Function
Sub AddNumNote()
'如果已经存在编号说明则不允许添加
 Dim ExistingNums1 As New VBA.Collection
        Set ExistingNums1 = colSearchName("MyNumNote")
        If ExistingNums1.Count <> 0 Then
            MsgBox "编号注释已经存在并已刷新！" & vbCrLf & " 不允许重复添加！", vbInformation, "臭豆腐工具箱|尺寸编号"
            Exit Sub
        End If


Set oSel = oActDoc.Selection
oSel.Clear

Dim DrawingView
Dim DrawingTexts
Dim i As Integer
For i = 1 To oActDoc.Sheets.ActiveSheet.Views.Count
    Set DrawingView = oActDoc.Sheets.ActiveSheet.Views.Item(i)
    If DrawingView.ViewType = catViewMain Then
    DrawingView.Activate
    Exit For
    End If
Next

 Set DrawingTexts = DrawingView.Texts
 Dim DrawingWindowLocation(1)
 Dim Status
 Status = oActDoc.Indicate2D("鼠标单击放置编号说明, Esc 退出", DrawingWindowLocation)
 If (Status = "Cancel") Then Exit Sub
 Dim sNotes As String
 Dim NumNotes As Object
 Dim n1, n2 As Integer
 Dim s2 As String
 n1 = 0
 n2 = 0
 s2 = ""
 sNotes = "#代表尺寸编号, 共有编号 " & n1 & "个, " & "最大的编号是 " & n2 & vbCrLf & Space(7) & "缺失的编号是: " & s2
 Set NumNotes = DrawingTexts.Add(sNotes, DrawingWindowLocation(0), DrawingWindowLocation(1))
 NumNotes.Name = "MyNumNote"
' FormatDimTag ManuNo
 
'txtNextno.Value = txtNextno.Value + 1
Set DrawingView = Nothing
Set DrawingTexts = Nothing
Set NumNotes = Nothing

End Sub


Attribute VB_Name = "A1l_CDF_Tool_2D3D"
Option Explicit
Public CATIA As Object, oActDoc As Object, objExcel As Object
Public oSel As Object
'Const CATVBAPATH = "D:\UGmeetsCATIA\CDF_Toolkit_CATIA"

Sub CATMain()
IntCATIA

CDF_Tool.Show vbModeless
CDF_Tool.Left = CDF_Tool.Left * 2
End Sub

Public Function oCATVBA_Folder(Optional ByVal sFolder As String)
 '取得于当前CATVBA文件路径同目录的子文件夹,无参数sFolder则取得当前catvba所在目录，
 '有参数sFolder则取得或创建与.catvba同一父目录的子目录sFolder

Dim oFSO, CATVBA_Dir
   Set oFSO = CreateObject("Scripting.FileSystemObject")
'   CATVBA_Dir = oFSO.GetParentFolderName(CATVBA_Path)
   CATVBA_Dir = "D:\Coding\Myhub\Macro_menu"
If sFolder <> "" Then
        CATVBA_Dir = CATVBA_Dir & "\" & sFolder
        'MsgBox CATVBA_Dir
            If Not oFSO.FolderExists(CATVBA_Dir) Then
                oFSO.CreateFolder (CATVBA_Dir)
            End If
End If
    
Set oCATVBA_Folder = oFSO.GetFolder(CATVBA_Dir)

End Function

Public Sub IntCATIA()
    On Error Resume Next
        Set CATIA = GetObject(CATIA.ActiveDocument.FullName, "CATIA.Application")
        If Err.Number <> 0 Then
            Set CATIA = CreateObject("CATIA.Application")
        End If
  
    Set oActDoc = CATIA.ActiveDocument
    On Error GoTo 0
End Sub
Public Sub IntExcel()
    On Error Resume Next
        Set objExcel = GetObject(, "Excel.Application")
        If Err.Number <> 0 Then
            Set objExcel = CreateObject("Excel.Application")
            
        End If
    On Error GoTo 0
End Sub
Public Function DwgLinkedDoc(oDocDwg)
On Error Resume Next
    If (TypeName(oDocDwg) <> "DrawingDocument") Then
    Exit Function
    End If
    
Dim DrawingSheets1, drawingSheet1, drawingViews1, drawingViewGenerativeBehavior1

    Set DrawingSheets1 = oDocDwg.Sheets
    Set drawingSheet1 = DrawingSheets1.ActiveSheet
    Set drawingViews1 = drawingSheet1.Views
    
    Dim i As Integer
    Dim Flag As Boolean
    Flag = False
    ' This VBA Macro Developed by Charles.Tang
    ' WeChat Chtang80,CopyRight reserved
    Dim FrontViewTmp As DrawingView
    Set FrontViewTmp = drawingViews1.Item(1)
    'catViewBackground\0,  catViewFront\1,  catViewLeft\2,  catViewRight\3,  catViewTop\4,  catViewBottom\5,  catViewRear\6,  catViewAuxiliary\7,  catViewIsom\8
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewFront Then  '
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewLeft Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewTop Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewIsom Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewRight Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewBottom Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If

    If Flag = False Then
        For i = 1 To drawingViews1.Count
            If drawingViews1.Item(i).ViewType = catViewRear Then
            Set FrontViewTmp = drawingViews1.Item(i)
                FrontViewTmp.Activate
                Set drawingViewGenerativeBehavior1 = FrontViewTmp.GenerativeBehavior
                Set DwgLinkedDoc = drawingViewGenerativeBehavior1.Document.Parent
                    If Err.Number = 0 Then
                        Flag = True
                        Exit For
                    End If
            End If
        Next
    End If
    
    
    Set DrawingSheets1 = Nothing
    Set drawingSheet1 = Nothing
    Set drawingViews1 = Nothing
    Set FrontViewTmp = Nothing
    Set drawingViewGenerativeBehavior1 = Nothing
On Error GoTo 0
End Function

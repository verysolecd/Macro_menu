VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Change_Link 
   Caption         =   "臭豆腐工具箱CATIA版 | 更改数模链接"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "Change_Link.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Change_Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRunChange_Click()
If lblChangeLink1.Caption = "" Then
MsgBox "请选择目标数模！", , "更改数模链接"
Exit Sub
End If

Dim DrawingSheets1, drawingSheet1, drawingViews1, drawingViewGenerativeBehavior1, drawingViewGenerativeLinks1

    Set DrawingSheets1 = oActDoc.Sheets
    Set drawingSheet1 = DrawingSheets1.ActiveSheet
    Set drawingViews1 = drawingSheet1.Views
    
    Dim i As Integer
    Dim Flag As Boolean
    Flag = False
    
    Dim oAimDoc As Object
    Set oAimDoc = CATIA.Documents.Read(lblChangeLink1.Caption)

 For i = 1 To drawingViews1.Count
    On Error Resume Next
    Set drawingViewGenerativeBehavior1 = drawingViews1.Item(i).GenerativeBehavior
    Set drawingViewGenerativeLinks1 = drawingViews1.Item(i).GenerativeLinks

    drawingViewGenerativeLinks1.RemoveAllLinks
    Select Case drawingViews1.Item(i).ViewType
           Case 0 '"catViewBackground"
           Case 12 '"catViewMain"
           Case 13 '"catViewPure_Sketch"
           Case 14 '"catViewUntyped"
           Case Else
           drawingViewGenerativeBehavior1.Document = oAimDoc.Product
    End Select
    ' This VBA Macro Developed by Charles.Tang
    ' WeChat Chtang80,CopyRight reserved
    
 Next

oActDoc.Update
GB_Frame.RefreshTitleBlock oActDoc
oActDoc.Update

MsgBox "更改完成，请按 Ctrl+S 保存！" & vbCrLf & vbCrLf & _
       "#重要提醒#" & vbCrLf & _
        "#请检查尺寸关联是否仍然有效！#", , "更改数模链接完成"
Unload Change_Link
End Sub

Private Sub cmdSelLink_Click()

lblChangeLink1.Caption = CATIA.FileSelectionBox("Select a product/part file", "*.CATPart;*.CATProduct", CatFileSelectionModeOpen)

End Sub

Private Sub Label1_Click()

End Sub


Private Sub CurrFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Open_Current_Folder.CATMain
End Sub

Private Sub UserForm_Initialize()
IntCATIA
    If (TypeName(oActDoc) <> "DrawingDocument") Then
        MsgBox "此命令只能在工程制图模块下运行", vbInformation, "Information"
        Exit Sub
    End If
On Error Resume Next
lblOrigLink1.Caption = DwgLinkedDoc(oActDoc).FullName
End Sub

Attribute VB_Name = "ASM_LocalSave"

Option Explicit
Private docs As Object

Sub CATMain()
    Dim origAlert As Boolean: origAlert = CATIA.DisplayFileAlerts
    CATIA.DisplayFileAlerts = False
    If Not CanExecute("ProductDocument") Then GoTo Cleanup
    Dim savePath As String: savePath = KCL.selFdl
    If savePath = "" Then GoTo Cleanup
    On Error GoTo ErrHandler
    Dim rootProd As Product: Set rootProd = CATIA.ActiveDocument.Product
    Set docs = KCL.InitDic   ' 初始化字典，仅一次
    Dim maxLvl As Integer: maxLvl = 0
    Call recurTreeLV(1, rootProd, docs, maxLvl)
    Call SaveByLV(docs, maxLvl, savePath)
    MsgBox "批量保存完成！" & vbCrLf & "保存路径：" & savePath, vbInformation, "CATIA批量保存"
Cleanup:
    CATIA.DisplayFileAlerts = origAlert
    Set docs = Nothing
    Set rootProd = Nothing
    Exit Sub
ErrHandler:
    MsgBox "保存失败：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical, "CATIA批量保存"
    Resume Cleanup
End Sub

' 递归遍历装配体，使用数组 (level, product) 作为字典值
Sub recurTreeLV(ByVal lvl As Integer, ByRef aProd As Product, ByRef dict As Object, ByRef maxLvl As Integer)
    If lvl > maxLvl Then maxLvl = lvl
    Dim pn As String: pn = Trim(aProd.partNumber)
    If pn = "" Then pn = "Unnamed_" & Replace(CreateObject("Scriptlet.TypeLib").GUID, "-", "")
    If Not dict.Exists(pn) Then
        dict.Add pn, Array(lvl, aProd) ' 第0项层级，第1项产品对象
    End If
    Dim i As Integer
    For i = 1 To aProd.Products.count
        Call recurTreeLV(lvl + 1, aProd.Products.item(i), dict, maxLvl)
    Next i
End Sub

' 按层级从深到浅保存，读取数组中的信息
Sub SaveByLV(ByRef dict As Object, ByVal maxLvl As Integer, ByVal folder As String)
    Dim lvl As Integer, key As Variant, info As Variant, suffix As String, target As Document, fullPath As String, i
    For lvl = maxLvl To 1 Step -1
        For Each key In dict.keys
            info = dict(key) ' info(0)=level, info(1)=product
            If info(0) = lvl Then
                Dim prod As Product: Set prod = info(1)
                Select Case TypeName(prod.ReferenceProduct.Parent)
                    Case "PartDocument": suffix = ".CATPart"
                    Case "ProductDocument":
                            Select Case info(0)
                                Case 1: suffix = ".CATProduct"
                                Case Else
                                    Dim str1: str1 = info(1).ReferenceProduct.Parent.FullName
                                    Dim Ary: Ary = Split(str1, ".")
                                         For i = LBound(Ary) To UBound(Ary)
                                              If info(1).ReferenceProduct.partNumber = Ary(i) Then suffix = ".CATProduct"
                                         Next i
                            End Select
                End Select
                If suffix <> "" Then
                    fullPath = folder & "\" & key & suffix
                    Set target = prod.ReferenceProduct.Parent
                    If target.FullName <> fullPath Then target.SaveAs fullPath
                End If
                dict.Remove key
            End If
        Next key
    Next lvl
End Sub

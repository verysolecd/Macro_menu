Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor:16744703}
Sub initme()
    MsgBox "你按了m42"
    Set allpn = KCL.InitDic(vbTextCompare)

    dim pdm: set pdm=new Class_PDM
    if not gprd is nothing then
        ini(oprd,pdm)        
        else
            MsgBox "请先选择要初始化的产品"
    end if

End Sub


sub ini(oprd,pdm)
    If allpn.exists(oPrd.PartNumber)=false Then
        allPN(oPrd.PartNumber) = 1
        Call initprd(oPrd)
    End If
    For Each product In oPrd.Products
        Call iniPrd(product, oDict)
    Next 
    allPN.RemoveAll
end sub

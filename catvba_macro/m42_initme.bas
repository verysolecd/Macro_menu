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
    set pdm=new Class_PDM
    if not gprd is nothing then
        dim oPrd
        set oPrd=gprd
        If allpn.exists(oPrd.PartNumber)=false Then
            allPN(oPrd.PartNumber) = 1
            Call pdm.initprd(oPrd)
        end if
            For Each product In oPrd.Products
                Call pdm.iniPrd(product)        
        Next 
            allPN.RemoveAll     
    else
            MsgBox "请先选择要初始化的产品"
    end if
End Sub

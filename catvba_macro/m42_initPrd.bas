Attribute VB_Name = "m42_initPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:initme}
'{Caption:初始化模板}
'{ControlTipText:将选择的产品和子产品文档按模板格式化}
'{BackColor:16744703}
Sub initme()

MsgBox "你按了m42"

dim pdm: set pdm=new Class_PDM

if not gprd is not nothing then
    pdm.initPrd gprd
else
    MsgBox "请选择产品"
end if
End Sub


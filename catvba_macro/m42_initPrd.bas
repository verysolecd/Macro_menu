Attribute VB_Name = "m42_initPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:initme}
'{Caption:��ʼ��ģ��}
'{ControlTipText:��ѡ��Ĳ�Ʒ���Ӳ�Ʒ�ĵ���ģ���ʽ��}
'{BackColor:16744703}
Sub initme()

MsgBox "�㰴��m42"

dim pdm: set pdm=new Class_PDM

if not gprd is not nothing then
    pdm.initPrd gprd
else
    MsgBox "��ѡ���Ʒ"
end if
End Sub


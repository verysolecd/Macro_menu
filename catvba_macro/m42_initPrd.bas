Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:initme}
'{Caption:��ʼ��ģ��}
'{ControlTipText:��ѡ��Ĳ�Ʒ���Ӳ�Ʒ�ĵ���ģ���ʽ��}
'{BackColor:16744703}
Sub initme()
    MsgBox "�㰴��m42"
    Set allpn = KCL.InitDic(vbTextCompare)
    set pdm=new Class_PDM
    if not gprd is nothing then
        If allpn.exists(oPrd.PartNumber)=false Then
            allPN(oPrd.PartNumber) = 1
            Call initprd(oPrd)
        end if
            For Each product In oPrd.Products
                Call iniPrd(product)        
        Next 
            allPN.RemoveAll     
    else
            MsgBox "����ѡ��Ҫ��ʼ���Ĳ�Ʒ"
    end if
End Sub

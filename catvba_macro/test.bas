Attribute VB_Name = "test"
Sub CATMain()



Dim xlm
Set xlm = New Class_XLM

xlm.init
xlm.inject_data 2, Array("P001", "2", 10, 1, "Part 1", "Image 1", 10, 1, 1, 1), "bom"
xlm.lvmg

'
'Set pdm = New class_PDM
'pdm.init
'pdm.selPrd



End Sub

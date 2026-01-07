Sub CATMain()
Set oPartDocument = CATIA.Documents
Set oPart = oPartDocument.Add("Part")
Set oPartDocument = CATIA.ActiveDocument
Set oPart = oPartDocument.Part
Set Myproduct = CATIA.ActiveDocument.GetItem(oPart.name)
Set oSelection = oPartDocument.Selection
Set MyVisProperties = oSelection.VisProperties 
Myproduct.PartNumber = "group_serial num_project no" 
'Set bodies1 = oPart.Bodies
Set oBody = oPart.MainBody
Dim oHBodies 'As HybridBodies
Set oHBodies = oPart.HybridBodies
Dim oHBody 'As HybridBody
'重命名主实体
oBody.name = "Part_Body"
'给实体随机上色
oselection.Clear()
oselection.add(oBody)
Randomize
'色调 0°~360、饱和度0~1、明度0~1 
'MyH = 60+rnd*240
MyH = (rnd*rnd*10000 mod 280) + 40   '从40~280的色相选色，避开红色

Randomize
MyS = 0.518+rnd*0.2
Randomize
MyV = 0.518+rnd*0.2
dim MyRGB
MyRGB = HSV2RGB(MyH,MyS,MyV)
R = MyRGB(1)
R = int (R)
G = MyRGB(2)
G = int(G)
B = MyRGB(3)
B = int(B)
'Msgbox("R:" & R & " G:" & G  & " B:" & B)
'Msgbox("H:" & MyH/1.5 & " S:" & int(MyS*240)  & " V:" & int(MyV*240))
MyVisProperties.SetRealColor R ,G ,B ,1
'删除所有几何图形集
oSelection.Clear()
if oHBodies.count>0 then
    for i = 1 to oHBodies.count
        Set oHybridBody = oHBodies.Item(i)
        oSelection.Add(oHybridBody)
    Next
oSelection.delete
end if
'增加#草图模板_sketch几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#草图模板"
'增加#Final_Design几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Final_Design_最终结果"
'增加#Input information几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Input information_输入信息"
'增加#Input information里面的Styling_Surfac几何图形集
Dim Input_information_HBodies 'As HybridBodies
set Input_information_HBodies=oHybridBody.HybridBodies
Set oHybridBody = Input_information_HBodies.Add()
oHybridBody.Name = "#Styling_Surface_造型面"
'增加#Input information里面的#Imported_Geometry几何图形集
Set oHybridBody = Input_information_HBodies.Add()
oHybridBody.Name = "#Imported_Geometry_输入元素"
'增加#Input information里面的#Section几何图形集
Set oHybridBody = Input_information_HBodies.Add()
oHybridBody.Name = "#Section_断面"
'增加#Basic_Surface几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Basic_Surface_基础结构"
'增加#Trimming几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Spilt_修边"
'增加#Spilt/Holes几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Holes_开孔"
'增加#Manufacturing_information几何图形集
Set oHybridBody = oHBodies.Add()
oHybridBody.Name = "#Manufacturing_information_工艺信息"
'增加#Manufacturing_information里面的RPS几何图形集
Dim Manufacturing_information_HBodies 'As HybridBodies
set Manufacturing_information_HBodies=oHybridBody.HybridBodies
Set oHybridBody = Manufacturing_information_HBodies.Add()
oHybridBody.Name = "#RPS"
'增加#Manufacturing_information里面的#Imported_Geometry几何图形集
Set oHybridBody = Manufacturing_information_HBodies.Add()
oHybridBody.Name = "#Matching_Areas_贴合面"

'设置线性与颜色
oselection.Clear()
oselection.add(oHybridBody)
MyVisProperties.SetRealColor 255 ,0 ,255 ,1
MyVisProperties.SetRealLineType 5,1
'增加#Manufacturing_information里面的#Section几何图形集
Set oHybridBody = Manufacturing_information_HBodies.Add()
oHybridBody.Name = "#Hoel_Function_孔位描述"
'增加#Manufacturing_information里面的#包边基准线几何图形集
Set oHybridBody = Manufacturing_information_HBodies.Add()
oHybridBody.Name = "#包边基准线"
'隐藏坐标系
oSelection.Clear()
Set MyXY=oPart.originElements.PlaneXY
Set MyYZ=oPart.originElements.PlaneYZ
Set MyZX=oPart.originElements.PlaneZX
oSelection.Add(MyXY)
oSelection.Add(MyYZ)
oSelection.Add(MyZX)
MyVisProperties.SetShow catVisPropertyNoShowAttr
oSelection.Clear()
oPart.Update 
'msgbox "模型树已构建完成，注意更改零件编号"
End Sub
function HSV2RGB(H,S,V)
'色调 0°~360、饱和度0~1、明度0~1
Dim C
Dim X
Dim m
Dim RGB(3)  '存储RGB的数组
hi = (H /60)  mod 6
f = h/60 - hi
p = V * (1-s)
q = V * (1-f*s)
t = V * (1-(1-f)*s)
'Msgbox("hi:" & hi & " f:" & f  & " p:" & p & " q:" & q & " q:" & q & " t:" & t)
if Hi = 0 then 
    RGB(1) = V
    RGB(2) = t
    RGB(3) = p
end if
if Hi = 1 then 
    RGB(1) = q
    RGB(2) = v
    RGB(3) = p
end if
if Hi = 2 then 
    RGB(1) = p
    RGB(2) = v
    RGB(3) = t
end if

if Hi = 3 then 
    RGB(1) = p
    RGB(2) = q
    RGB(3) = v
end if

if Hi = 4 then 
    RGB(1) = t
    RGB(2) = p
    RGB(3) = v
end if
if Hi = 5 then 
    RGB(1) = V
    RGB(2) = p
    RGB(3) = q
end if
'Msgbox("R:" & RGB(1) & " G:" & RGB(2)  & " B:" & RGB(3) )
RGB(1) = RGB(1) *255
RGB(2) = RGB(2) *255
RGB(3) = RGB(3) *255
HSV2RGB = RGB

end function 














# BOM 生成重构方案 (最终版)

本方案将 [RW_Cbom.bas](file:///d:/catia/Macro_menu/Src/RW_Cbom.bas)、`Cls_PDM` 和 `Cls_XLM` 的 BOM 数据流重构为基于结构体（Type）的模式，并将所有公共 Type 定义在 [A00_globalVar.bas](file:///d:/catia/Macro_menu/Src/A00_globalVar.bas) 中。同时废弃 `Cls_Para` 类。

## 核心变更点
1.  **数据中心化**: [A00_globalVar.bas](file:///d:/catia/Macro_menu/Src/A00_globalVar.bas) 承载 `BOMItem` 和 `ParamItem` 定义。
2.  **移除冗余**: 删除 [Cls_Para.cls](file:///d:/catia/Macro_menu/Src/Cls_Para.cls)。
3.  **类型安全**: `Cls_PDM` 产出 Type 数组，`Cls_XLM` 消费 Type 数组。

## 详细步骤

### 1. [修改] [A00_globalVar.bas](file:///d:/catia/Macro_menu/Src/A00_globalVar.bas)
在文件头部添加全局 Type 定义：
```vb
Public Type BOMItem
    Level As Integer        ' 层级
    PartNumber As String    ' 件号
    Nomenclature As String  ' 英文名称
    Definition As String    ' 中文名称
    InstanceName As String  ' 实例名
    Quantity As Long        ' 数量
    Mass As Double          ' 单重
    Material As String      ' 材质
    Thickness As Double     ' 厚度 AB对应10
    Density As Double       ' 密度
    TotalMass As Double     ' 总重
    ' 预留扩展
    UserProp1 As String
    UserProp2 As String
End Type

Public Type ParamItem
    Name As String
    ParamType As String
    Value As Variant
    Target As Object    ' 指向 CATIA Parameter 对象
    Description As String
End Type
```

### 2. [删除] [Cls_Para.cls](file:///d:/catia/Macro_menu/Src/Cls_Para.cls)
该类功能已完全由 `ParamItem` 替代，直接从项目中移除。

### 3. [修改] [Cls_PDM.cls](file:///d:/catia/Macro_menu/Src/Cls_PDM.cls)
- **替换 Cls_Para**: 将所有 `New Cls_Para` 的代码替换为 `Dim p As ParamItem`。
- **重构 infoPrd**:
    ```vb
    Public Function infoPrd(oPrd As Object) As BOMItem
        Dim item As BOMItem
        ' ... 赋值逻辑 ...
        infoPrd = item
    End Function
    ```
- **重构 ProduceBOM**: 返回 `BOMItem()` 数组。

### 4. [修改] [Cls_XLM.cls](file:///d:/catia/Macro_menu/Src/Cls_XLM.cls)
- **新增 InjectBOM_Typed(data() As BOMItem)**:
    - 内部将 `BOMItem` 数组转换为二维 Variant 数组。
    - 一次性写入 Excel。
    - **列映射逻辑**: 在转换过程中，显式指定 `Arr(i, 3) = item.Nomenclature`。

### 5. [修改] [RW_Cbom.bas](file:///d:/catia/Macro_menu/Src/RW_Cbom.bas)
- 协调调用 `pdm.ProduceBOM` 和 `xlm.InjectBOM_Typed`。

## 立即执行
如果确认无误，我将按照此方案开始编写代码：
1. 修改 [A00_globalVar.bas](file:///d:/catia/Macro_menu/Src/A00_globalVar.bas)
2. 修改 [Cls_PDM.cls](file:///d:/catia/Macro_menu/Src/Cls_PDM.cls)
3. 修改 [Cls_XLM.cls](file:///d:/catia/Macro_menu/Src/Cls_XLM.cls)
4. 修改 [RW_Cbom.bas](file:///d:/catia/Macro_menu/Src/RW_Cbom.bas)
5. 删除 [Cls_Para.cls](file:///d:/catia/Macro_menu/Src/Cls_Para.cls)


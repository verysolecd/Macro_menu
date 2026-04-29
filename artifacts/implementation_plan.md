# 解耦 cls_MnUI 与 Cls_DynaUIEngine

## 背景与问题

这两个类当前承担了**完全不同的职责**，却存在直接的代码依赖：

| 类 | 本职 |
|----|------|
| `cls_MnUI` | **菜单项数据模型** — 解析单个模块的 `{GP:}{Ep:}` 标签，保存元数据，执行宏 |
| `Cls_DynaUIEngine` | **动态弹窗引擎** — 解析 `%UI` 注释，动态生成 UserForm 控件，收集结果 |

两者理论上**完全独立**：一个属于"主菜单扫描/执行"链路，一个属于"子宏弹窗"链路。但当前代码中存在 **3 处耦合**，导致无法独立使用或修改任意一方。

---

## 耦合点详细分析

### 耦合点 1 — `Cls_DynaUIEngine` 直接 `New cls_MnUI`（最严重）

**位置**：`Cls_DynaUIEngine.cls` 第 162、176、409、416 行

```vba
' LoadFromMenuTags 方法内部
Dim menuItem As cls_MnUI              ' ← 强类型引用 cls_MnUI
Set menuItem = TryParseMenuModule(...)
...
Dim item As cls_MnUI                  ' ← 强类型引用

' TryParseMenuModule 私有函数
Private Function TryParseMenuModule(...) As cls_MnUI  ' ← 返回值是 cls_MnUI
    Dim menuItem As New cls_MnUI      ' ← 在引擎内部创建数据模型对象
    If Not menuItem.InitFromCode(...) Then Exit Function
    ...
    Set TryParseMenuModule = menuItem
End Function
```

**问题**：`Cls_DynaUIEngine` 是弹窗引擎，却在内部创建并操作菜单扫描专用的数据模型 `cls_MnUI`。  
`LoadFromMenuTags` 是主菜单扫描功能，**根本不应该存在于弹窗引擎中**，这两个功能没有任何关联。

---

### 耦合点 2 — `Cls_DynaUIEngine` 硬编码 `New CAT_springWD`

**位置**：`Cls_DynaUIEngine.cls` 第 240 行

```vba
Set mWD = New CAT_springWD   ' ← 具体 UserForm 类型被硬编码在引擎内
```

**问题**：引擎与具体的窗体模板强绑定，无法替换其他 UserForm，也无法在不加载 `CAT_springWD` 的情况下单独测试引擎逻辑。（此处属于轻耦合，本次一并消除）

---

### 耦合点 3 — `cls_MnUI.ToDictionary()` 依赖 `KCL.InitDic`

**位置**：`cls_MnUI.cls` 第 143 行

```vba
Public Function ToDictionary() As Object
    Dim dic As Object
    Set dic = KCL.InitDic(vbTextCompare)   ' ← 数据模型依赖全局工具模块
```

**问题**：`cls_MnUI` 作为纯数据模型类，理论上不应依赖 `KCL` 全局模块。`KCL.InitDic` 仅是 `CreateObject("Scripting.Dictionary")` 的简单包装，完全可以内联。

---

## 解耦方案

> [!IMPORTANT]
> 核心原则：**最小改动，不破坏现有调用链**。`A00_Menu.bas` 和所有业务宏对外接口保持不变。

### 方案：将 `LoadFromMenuTags` 迁移回 `A00_Menu`（或独立工具函数）

`LoadFromMenuTags` 本质上是**主菜单扫描逻辑**，与弹窗引擎无关，应归属于 `A00_Menu`（或独立出去）。

**迁移后职责变更**：

```
修改前:
  A00_Menu → Cls_DynaUIEngine.LoadFromMenuTags → New cls_MnUI

修改后:
  A00_Menu → (自身扫描，或通过独立帮助函数) → New cls_MnUI
  Cls_DynaUIEngine → 只负责弹窗 UI，不再知道 cls_MnUI 的存在
```

---

## 具体修改内容

### ① `Cls_DynaUIEngine.cls` — 删除菜单扫描相关代码

#### [MODIFY] [Cls_DynaUIEngine.cls](file:///d:/catia/Macro_menu/Src/Cls_DynaUIEngine.cls)

**删除以下内容（约 50 行）**：
- `LoadFromMenuTags` 公开方法（整个方法体）
- `TryParseMenuModule` 私有方法（整个方法体）
- `SortDictListByModule` 私有方法（整个方法体）
- `Dim item As cls_MnUI` / `Dim menuItem As cls_MnUI` 等所有强类型变量声明
- 对应的 `TAG_MDLNAME` 常量（若仅被上述代码使用）

**保留全部弹窗逻辑**（`LoadFromModuleName`、`Show`、`ShowToolbar`、`Alert`、`BindEvents` 等）。

**消除 `New CAT_springWD` 硬编码**：通过工厂函数间接创建：
```vba
' 修改前
Set mWD = New CAT_springWD

' 修改后（仍用 New，但提取为独立私有函数，便于未来替换）
Set mWD = P_CreateSpringWD()

Private Function P_CreateSpringWD() As Object
    Set P_CreateSpringWD = New CAT_springWD
End Function
```
> 此步骤可选，视用户需求决定是否纳入本次改动。

---

### ② `A00_Menu.bas` — 接收迁移来的扫描逻辑

当前 `A00_Menu.bas` 已有自己完整的扫描逻辑（`GetMenuItems` + `ProcessModule` + `OrganizeForView`），与 `Cls_DynaUIEngine.LoadFromMenuTags` 做的是**完全相同的事情**（重复代码）。

迁移后：`A00_Menu` 继续使用自己现有的扫描逻辑，**无需变更**。

---

### ③ `cls_MnUI.cls` — 消除对 `KCL` 的依赖

#### [MODIFY] [cls_MnUI.cls](file:///d:/catia/Macro_menu/Src/cls_MnUI.cls)

**修改 `ToDictionary()` 方法**：
```vba
' 修改前
Set dic = KCL.InitDic(vbTextCompare)

' 修改后（直接内联，去掉对 KCL 的依赖）
Set dic = CreateObject("Scripting.Dictionary")
dic.CompareMode = vbTextCompare
```

---

## 修改范围汇总

| 文件 | 操作 | 行数变化 |
|------|------|---------|
| `Cls_DynaUIEngine.cls` | 删除 `LoadFromMenuTags`、`TryParseMenuModule`、`SortDictListByModule` 三个方法及相关类型引用 | -约50行 |
| `cls_MnUI.cls` | `ToDictionary()` 内 `KCL.InitDic` 改为内联 `CreateObject` | -1行 +2行 |
| `A00_Menu.bas` | **不需要改动**（已有完整扫描逻辑） |

---

## 影响评估

> [!WARNING]
> 如果项目中**其他地方调用了 `Cls_DynaUIEngine.LoadFromMenuTags`**，需要将调用方改为直接使用 `A00_Menu` 的现有扫描逻辑，或将该函数复制到 `A00_Menu`。

当前检查：`A00_Menu.bas` 调用的是自己的 `GetMenuItems()`，**未发现对 `LoadFromMenuTags` 的外部调用**（`A0TEST_Engine.bas` 有一行 `Test5_MainMenu` 测试，需确认）。

> [!NOTE]
> 解耦后，两条链路完全独立：
> - **主菜单链路**：`A00_Menu` → `cls_MnUI` → `Cat_Macro_Menu_View`
> - **弹窗链路**：业务宏 → `KCL.newEngine()` → `Cls_DynaUIEngine` → `CAT_springWD`

---

## 确认信息（已核查）

**`A0TEST_Engine.bas` 第 192 行有实际调用 `LoadFromMenuTags`：**
```vba
Set SoLst = oEng.LoadFromMenuTags(PageMap)   ' A0TEST_Engine.bas:192
```

因此解耦策略调整为：
- 将 `LoadFromMenuTags` 从 `Cls_DynaUIEngine` **迁移**到一个独立的位置（`A00_Menu.bas` 中的私有工具函数，或作为 `cls_MnUI` 的静态工厂函数），同时在 `Cls_DynaUIEngine` 中保留一个**转发包装**（或直接在 `A0TEST_Engine` 中改调 `A00_Menu` 的同名函数）。
- 最简方案：**`A0TEST_Engine` 改为调用 `A00_Menu` 模块级函数**（因为 `A00_Menu` 已有等价的 `GetMenuItems+OrganizeForView`），`Cls_DynaUIEngine` 内不再有任何 `cls_MnUI` 引用。

**`CAT_springWD` 硬编码**：本次一并提取为工厂函数，影响极小。

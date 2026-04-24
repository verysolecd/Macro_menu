# Macro_menu 项目架构分析

## 项目概述

这是一个运行在 **CATIA V5** 环境中的 **VBA 宏菜单管理系统**。  
核心思路：以"标签驱动"方式将分散的宏模块自动扫描、分组并呈现为一个统一的浮动菜单面板，用户点击按钮即可执行对应的 CATIA 操作。

---

## 整体架构分层

```
┌─────────────────────────────────────────────────────────┐
│                     主入口 A00_Menu                      │
│  扫描当前 VBA 项目 → 读取模块标签 → 组装并显示菜单      │
└────────────────────┬────────────────────────────────────┘
                     │
       ┌─────────────┼──────────────────────┐
       ▼             ▼                      ▼
┌────────────┐ ┌──────────────┐   ┌─────────────────────┐
│ 视图层 View│ │ 数据模型层   │   │  动态UI引擎层        │
│            │ │              │   │                      │
│Cat_Macro_  │ │ cls_MnUI     │   │ Cls_DynaUIEngine     │
│Menu_View   │ │ (菜单项数据) │   │ (动态生成任意表单)   │
│(主菜单面板)│ │              │   │                      │
│            │ │ Cls_MnuBtn   │   │ CAT_springWD         │
│VbaModule   │ │ EVT          │   │ (弹窗模板 UserForm)  │
│Maneger     │ │ (按钮事件)   │   └─────────────────────┘
│View (模块  │ └──────────────┘
│管理界面)   │
└────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│                   基础支撑层 (Framework)                  │
│  KCL.bas            - 全局工具库 / 公共变量              │
│  Cls_VbaUltiliseLib - 文件IO / Debug / 注册表工具        │
│  Cls_JsonConverter  - JSON 序列化/反序列化               │
│  Cls_PDM            - CATIA 产品文档监听 / BOM 处理      │
│  Cls_XLM            - Excel 交互封装                     │
│  Cls_VbaMdlMgr      - VBA 模块导入导出管理               │
└─────────────────────────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────────────────────────┐
│                  业务功能宏 (按分组 GP)                    │
│  GP=1  RW_*    - 项目流程（BOM / 修订 / 质量）           │
│  GP=3  ASM_*   - 装配体操作                              │
│  GP=4  MDL_*   - 零件 / 几何体操作                       │
│  GP=5  DRW_*   - 工程图操作                              │
│  GP=6  OTH_*   - 其他杂项工具                            │
│  GP=7  CAT_*   - CATIA 通用工具                          │
└─────────────────────────────────────────────────────────┘
```

---

## 核心机制：模块标签（Tag）驱动

每个业务宏模块在其声明区开头写注释形式的元数据标签：

```vba
Attribute VB_Name = "OTH_PrePn"
'{GP:6}                          ← 分组 ID（与 A00_Menu 中 GroupName 对应）
'{Ep:Pnmgr}                      ← 入口函数名（Entry Point）
'{Caption:PN管理}                 ← 按钮显示文字
'{ControlTipText:零部件编号管理}  ← 按钮提示文字
'{BackColor:}                     ← 按钮背景色（可选）
```

**扫描流程（A00_Menu → cls_MnUI）**：

1. `A00_Menu.CATMain()` 获取当前执行的 VBA 项目
2. 遍历所有 `Type=1`（标准模块）的 VBComponent
3. 读取各模块声明区代码，交由 `cls_MnUI.InitFromCode()` 用正则解析标签
4. 验证 `{GP:x}` 在预定义分组中、且入口函数存在
5. 将所有有效项按 GP 分组后传给 `Cat_Macro_Menu_View` 显示

---

## 文件清单与职责

### 🔴 主入口 & 配置
| 文件 | 职责 |
|------|------|
| `A00_Menu.bas` | 系统入口，扫描宏模块、组装菜单 |
| `A00_backmeup.bas` | 备份辅助入口 |

### 🟠 视图层（UserForm）
| 文件 | 职责 |
|------|------|
| `Cat_Macro_Menu_View.frm/.frx` | 主菜单面板：MultiPage 标签页 + 动态按钮 |
| `CAT_springWD.frm/.frx` | 轻量弹窗模板，由 DynaUIEngine 接管控件 |
| `VbaModuleManegerView.frm/.frx` | VBA 模块导入/导出管理界面 |
| `FrmFlower.frm/.frx` | Flower 特殊工具窗体 |

### 🟡 核心类库
| 文件 | 职责 |
|------|------|
| `cls_MnUI.cls` | 菜单项数据模型（解析标签、执行宏） |
| `Cls_MnuBtnEVT.cls` | 按钮事件绑定（脚本模式 / 结果收集模式） |
| `Cls_DynaUIEngine.cls` | **动态UI引擎**：解析 `%UI` 注释，自动生成表单控件 |
| `Cls_PDM.cls` | CATIA 产品文档事件监听、BOM 遍历、属性读写 |
| `Cls_VbaMdlMgr.cls` | VBA 模块的导入/导出、版本管理 |
| `Cls_VbaUltiliseLib.cls` | 通用工具（文件读写、调试输出、等待） |
| `Cls_JsonConverter.cls` | JSON 解析与序列化 |
| `Cls_XLM.cls` | Excel 自动化封装 |
| `Cls_WsEvt.cls` | 工作空间事件监听 |

### 🟢 全局基础模块
| 文件 | 职责 |
|------|------|
| `KCL.bas` | **全局枢纽**：公共变量、Win API 声明、通用函数、工厂方法（`newEngine`、`newFrm`、`InitDic` 等） |

### 🔵 业务功能宏（GP 分组）

#### GP:1 — R&W（项目流程）
| 文件 | 功能 |
|------|------|
| `RW_1setgprd.bas` | 设置 Group Product |
| `RW_2freegPrd.bas` | 释放 Group Product |
| `RW_3initme.bas` | 初始化产品 |
| `RW_6nosamebdy.bas` | 去除重复 Body |
| `RW_Cbom.bas` | 生成 BOM |
| `RW_Revise.bas` | 修订工具栏 |
| `RW_cMass.bas` | 质量计算 |

#### GP:3 — ASM（装配体）
| 文件 | 功能 |
|------|------|
| `ASM_1ex2stp.bas` | 导出 STP/ZIP |
| `ASM_2Localsend.bas` | 发送到本地目录 |
| `ASM_3LocalSave.bas` | 本地保存 |
| `ASM_CMP.bas` | 比较装配体 |
| `ASM_ChildMng.bas` | 子产品管理 |
| `ASM_NewBH.bas` | 新建 Body Header |
| `ASM_Updateme.bas` | 强制更新 |
| `ASM_reorderPrd.bas` | 重排产品顺序 |
| `ASM_weldSel.bas` | 焊缝选择创建 |

#### GP:4 — MDL（零件建模）
| 文件 | 功能 |
|------|------|
| `MDL_Bodyrename.bas` | Body 批量重命名 |
| `MDL_LayersMng.bas` | 图层管理 |
| `MDL_MaterialColors.bas` | 材料颜色应用 |
| `MDL_Part2Product.bas` | Part 转 Product |
| `MDL_addgeotree.bas` | 添加几何体集 |
| `MDL_addsubgeo.bas` | 添加子几何体 |
| `MDL_hasLeftAxis.bas` | 检查左手轴 |
| `MDL_holecenter.bas` | 孔中心点提取 |
| `MDL_pt2Hb.bas` | 点复制到 HybridBody |
| `MDL_pt2xl_abscoord.bas` | 点坐标导出 Excel |
| `MDL_rmCrv.bas` | 删除曲线 |
| `MDL_setThreadcolor.bas` | 螺纹着色 |
| `MDL_wfrename.bas` | Wire Frame 重命名 |
| `MDL_Shapeinfo.bas` | 形状信息提取 |

#### GP:5 — DRW（工程图）
| 文件 | 功能 |
|------|------|
| `DRW_DrwLock.bas` | 图纸锁定/解锁 |
| `DRW_ExPDF.bas` | 导出 PDF |
| `DRW_ShtPage.bas` | 图页管理 |
| `DRW_Tb2xl.bas` | 2D 表格导出 Excel |
| `DRW_VIewBOM.bas` | 视图 BOM 创建 |
| `DRW_newTol.bas` | 新建公差 |
| `DRW_viewBOM_drawing_template.bas` | BOM 模板 |
| `Drw_myframe.bas` / `Drw_myframe2.bas` | 图框生成 |

#### GP:6 — OTRS（其他工具）
| 文件 | 功能 |
|------|------|
| `OTH_PrePn.bas` | **PN 编号批量管理**（前缀/后缀/删除） |
| `OTH_3Dmark.bas` | 3D 标注 Label |
| `OTH_Flower.bas` | Flower 特殊工具 |
| `OTH_Minibox.bas` | 迷你弹窗工具 |
| `OTH_OPenRR.bas` | 打开最近记录 |
| `OTH_capture.bas` | 截图保存 |
| `OTH_designlog.bas` | 设计日志 |
| `OTH_ivhideshow.bas` | 显示/隐藏工具栏 |
| `OTH_unfoldme.bas` | 展开子产品 |

#### GP:7 — CATIA 通用
| 文件 | 功能 |
|------|------|
| `CAT_Color.bas` | 背景颜色切换 |
| `CAT_Filepath.bas` | 文件路径工具 |
| `CAT_SWScr.bas` | 屏幕刷新切换 |
| `CAT_closePartWindows.bas` | 关闭零件窗口 |

### ⚪ 其他
| 文件 | 功能 |
|------|------|
| `KCL.bas` | 全局库（如前述） |
| `A0TEST.bas` / `A0TEST_Engine.bas` | DynaUIEngine 测试用例 |
| `ZZ_BACK2.bas` / `ZZ_BCK.bas` / `zz_BCKcurve.bas` | 历史备份 |

---

## 动态 UI 引擎（Cls_DynaUIEngine）详解

这是项目中最重要的创新机制之一，**允许宏模块无需手动创建 UserForm，直接通过代码注释声明所需的 UI 控件**：

```vba
' %UI Label   lbl_info    提示文字
' %UI CheckBox chk_opt1   选项A
' %UI TextBox  txt_input  输入框
' %UI Button   btnOK      确定
' %UI Button   btnCancel  取消
```

引擎解析 `%UI <Type> <Name> <Caption>` 格式，动态将控件添加到 `CAT_springWD` 模板窗体上，并自动绑定事件。

**调用方式**（以 `OTH_PrePn.bas` 为例）：
```vba
Dim oEng As Object
Set oEng = KCL.newEngine(mdlname)  ' 读取本模块 %UI 注释，生成表单
oEng.Show                           ' 显示（模态）
Select Case oEng.ClickedButton
    Case "btnOK"
        istr = oEng.Results("txt_str")  ' 读取输入值
```

---

## 关键设计模式

| 模式 | 说明 |
|------|------|
| **标签驱动注册** | 模块无需手动注册，加标签即自动出现在菜单 |
| **工厂方法** | `KCL.newEngine()` / `KCL.newFrm()` 统一创建 UI 对象 |
| **事件代理** | `Cls_MnuBtnEVT` 用 `WithEvents` 延长按钮事件对象生命周期 |
| **Observer** | `Cls_PDM` 监听 CATIA 活动产品变化，实时更新菜单状态栏 |
| **JSON 持久化** | `Cls_VbaMdlMgr` 用 JSON 记录用户自定义导出路径 |

---

## 重要全局变量（KCL.bas）

| 变量 | 类型 | 用途 |
|------|------|------|
| `rootDoc` | Object | 当前根文档 |
| `rootPrd` | Object | 当前根产品 |
| `pdm` | Cls_PDM | 全局产品文档监听器 |
| `xlm` | Cls_XLM | 全局 Excel 操作对象 |
| `xlAPP` | Object | Excel 应用实例 |
| `g_allPN` | Object(Dictionary) | 全局零件编号缓存 |
| `g_Btn` | Object | 最后点击的按钮引用（传给脚本） |
| `g_Picpath` | Variant | 图片路径 |

---

## 注意事项（后续修改参考）

> [!NOTE]
> 中文注释在源文件中部分因编码问题显示为乱码（`?`），实际运行在 CATIA 的 GB2312 环境下正常。修改时注意文件编码。

> [!IMPORTANT]
> 新增业务宏只需：1) 创建 `.bas` 文件；2) 在声明区加 `{GP:x}` 等标签；3) 实现入口函数。无需改动主菜单代码。

> [!WARNING]
> `Cls_VbaMdlMgr.import_project()` 存在已知问题（参见历史会话：ce357595），导入 `.bas/.cls` 模块到 CATIA VBA 项目时可能失败。

> [!TIP]
> `Cls_DynaUIEngine` 的 `%UI` 解析逻辑在 `LoadFromModuleName()` 方法中，使用 `%UEI` 前缀（注意拼写）也有效，是早期兼容写法。

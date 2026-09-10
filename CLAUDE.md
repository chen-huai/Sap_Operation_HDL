# CLAUDE.md

> 最后更新：2026-05-19 | 请用中文回复，所有测试模块都放在 `test/` 中

## 项目概述

**Sap_Operation_HDL** — 基于 Python + PyQt5 的 SAP 自动化工具，提供 SAP 订单自动创建、数据处理、PDF 发票重命名、营收工时分配等功能，具有多选项卡 GUI 界面。

## 架构总览

```mermaid
graph TD
    subgraph GUI层
        A[Sap_Operate_CPS.py<br/>主入口 MyMainWindow] --> B[Sap_Operate_Ui.py<br/>Qt Designer 生成 UI]
        A --> C[Data_Table.py<br/>表格窗口]
        C --> D[Table_Ui.py<br/>表格 UI 代码]
        A --> TM[theme_manager_theme.py<br/>主题管理]
    end

    subgraph Mixin业务层
        A --> M1[main_window_ui_mixin.py<br/>UI 初始化与主题]
        A --> M2[config_mixin.py<br/>配置加载/GUI 数据]
        A --> M3[sap_order_mixin.py<br/>odmDataToSap 订单批量]
        A --> M4[odm_invoice_mixin.py<br/>ODM 数据/发票重命名]
        A --> M5[hour_mixin.py<br/>工时录入]
    end

    subgraph SAP服务层
        M3 --> S[sap/ 包]
        M4 --> S
        M5 --> S
        S --> S1[sap/session.py<br/>SapSession win32com 会话]
        S --> S2[sap/services/<br/>OrderService/InvoiceService/HourService]
        S --> S3[sap/transactions/<br/>order/invoice/hours 事务流]
        S --> S4[sap/models.py<br/>SapConfig/OrderData/RevenueData/SapResult]
        S --> S5[sap/rules.py<br/>业务规则]
    end

    subgraph 数据/PDF/营收
        M2 --> F[Get_Data.py<br/>数据读取与转换]
        M4 --> H[PDF_Parser_Utils.py<br/>PDF 读取 + 字段提取]
        M5 --> I[Revenue_Operate.py<br/>营收工时分配]
        A --> J[Excel_Field_Mapper.py<br/>字段映射]
    end

    subgraph 基础设施
        A --> L[Logger.py<br/>日志记录]
        A --> O[chicon.py<br/>图标资源]
        A --> RG[runtime_globals.py<br/>跨模块全局变量]
    end

    subgraph auto_updater模块
        A --> P[auto_updater/<br/>自动更新系统]
        P --> P1[github_client.py]
        P --> P2[download_manager.py]
        P --> P3[backup_utils.py]
        P --> P4[update_executor.py]
        P --> P5[ui/ 子模块<br/>更新 UI 组件]
    end

    subgraph 构建
        Q[build_with_pyinstaller.py]
        R[build_with_signing.py]
    end
```

## 模块索引

| 模块 | 路径 | 职责 |
|------|------|------|
| 主入口 | `Sap_Operate_CPS.py` | `MyMainWindow` 多 mixin 继承入口，事件绑定 + 应用启动 |
| UI Mixin | `main_window_ui_mixin.py` | UI 初始化、主题、状态栏、版本菜单 |
| 配置 Mixin | `config_mixin.py` | `getConfig` 加载 csv、`getGuiData` 读取 GUI 字段 |
| 订单 Mixin | `sap_order_mixin.py` | `odmDataToSap` 多 sheet Excel 批量创建订单、`orderUnlockOrLock` 锁/解锁 |
| ODM Mixin | `odm_invoice_mixin.py` | ODM 数据合并/拆分、PDF 发票/电子发票重命名 |
| 工时 Mixin | `hour_mixin.py` | `hourOperate` 工时数据批量录入 SAP |
| SAP 包 | `sap/` | SAP 自动化分层：session / services / transactions / models / rules |
| SAP 会话 | `sap/session.py` | `SapSession` win32com 连接封装 |
| SAP 服务 | `sap/services/` | `OrderService` / `InvoiceService` / `HourService`，业务级 API |
| SAP 事务 | `sap/transactions/` | `OrderTransaction` / `InvoiceTransaction` / `HourTransaction`，具体 SAP 屏幕操作 |
| SAP 模型 | `sap/models.py` | `SapConfig` / `OrderData` / `OrderItemData` / `RevenueData` / `DataBEntry` / `PlanCostEntry` / `HourData` / `SapResult` 等 dataclass |
| SAP 规则 | `sap/rules.py` | A2 物料拆分、Data A 客户判定、Plan Cost 阈值等业务规则 |
| 数据读取 | `Get_Data.py` | `Get_Data` 类，Excel/CSV 读取与多 sheet 取数 |
| 字段映射 | `Excel_Field_Mapper.py` | `ExcelFieldMapper`，多命名风格字段匹配 |
| PDF 解析 | `PDF_Parser_Utils.py` | PDF 读取 + 发票字段提取（公司名、金额、发票号），原 `PDF_Operate.py` 已并入 |
| 营收分配 | `Revenue_Operate.py` | `RevenueAllocator`，工时与营收分配计算 |
| 表格窗口 | `Data_Table.py` | `MyTableWindow`，数据表格展示 |
| 日志 | `Logger.py` | `Logger` 类，基于 pandas 的操作日志 |
| 主题管理 | `theme_manager_theme.py` | `ThemeManager`，应用主题切换 |
| 图标资源 | `chicon.py` | 内嵌图标 base64 数据 |
| 运行时全局 | `runtime_globals.py` | 跨模块共享变量（`configContent` / `myWin` / `today` 等） |
| 自动更新 | `auto_updater/` | 基于 GitHub Releases 的完整更新系统 |
| 构建脚本 | `build_with_pyinstaller.py` | PyInstaller 打包 |
| 签名构建 | `build_with_signing.py` | 带代码签名的打包 |

## 关键依赖

- **PyQt5** 5.15.11 — GUI 框架
- **pandas** 2.2.2 — 数据处理
- **win32com** — SAP GUI Scripting 自动化（仅 Windows）
- **pdfplumber / pdfminer.six / pypdfium2** — PDF 解析
- **openpyxl** — Excel 读写
- **chinese_calendar** — 中国节假日判断（营收模块）
- **PyInstaller** — 构建可执行文件

## 常用命令

```bash
# 运行主程序
python Sap_Operate_CPS.py

# 构建可执行文件
python build_with_pyinstaller.py

# 手动 PyInstaller
pyinstaller --onefile --windowed --clean --noconfirm --icon=Sap_Operate_Logo.ico Sap_Operate_CPS.py
```

## 数据流

1. **输入** → Excel/CSV 文件经 `Get_Data.py` 读取（多 sheet 走 `getExcelSheetsData`）
2. **映射** → `Excel_Field_Mapper.py` 统一多命名风格字段
3. **对象适配** → mixin 内 `_build_*` 方法将 DataFrame 行转为 `sap/models.py` 中的 dataclass
4. **SAP 操作** → `sap/services/` 暴露业务 API → `sap/transactions/` 执行屏幕操作 → `sap/session.py` 通过 win32com 与 SAP GUI 通信
5. **PDF 处理** → `PDF_Parser_Utils.py` 读取并提取发票字段，重命名落盘
6. **营收分配** → `Revenue_Operate.py` 按工时分配营收
7. **输出** → GUI textBrowser 展示 + Excel 日志（每次批量任务写一份 log.xlsx）

## SAP 调用约定

- 所有 SAP 操作统一走 `sap/` 包；旧 `Sap_Function.py` 适配器已删除
- 业务代码持有 `SapSession.connect()` 单例，并将其注入 `OrderService` / `InvoiceService` / `HourService`
- 操作结果统一返回 `SapResult`，用 `.success` / `.message` 判定，避免旧的 `dict['flag']` 风格
- 全程 try / finally 中调用 `sap_session.close()` 释放

## SAP 集成要求

- SAP GUI 已安装并运行
- Scripting 已在 SAP GUI 中启用
- 用户具有相应 SAP 权限

## 配置

- 桌面 `config/config_sap_CPS.csv` — 用户配置文件
- `auto_updater/config_constants.py` — 版本号与更新配置

## 文件命名约定

- `*_Ui.py` — Qt Designer 生成的 UI 代码（勿手动编辑）
- `*.ui` — Qt Designer 源文件
- `*.ico` — 应用图标
- `dist/` / `build/` — 构建产物（已 gitignore）

## 全局规范

- 所有回复使用中文
- 测试文件放在 `test/` 目录
- UI 文件由 Qt Designer 生成，不手动编辑 `*_Ui.py`
- 字段映射通过 `Excel_Field_Mapper.py` 的映射表维护

## 批量流程硬约束（基本原则，不可违反）

**主程序的批量创建/编辑流程绝不允许死循环。** 批量流程跑在主线程（靠
`QApplication.processEvents()` 刷 UI，无 QThread 承载 SAP 任务），一旦卡住：UI 冻结、
关窗无响应、SAP 会话不释放、log 停在半截，用户只能强杀进程且 SAP 侧可能留下半成品订单。

改动 `odmDataToSap` / `_edit_order_row` / `hourOperate` / `orderUnlockOrLock` 或
`sap/transactions/` 内任何循环时，交付前逐条自检：

1. **终止条件确定** — `while` 必须有硬上限。参考 `_sum_item_net_values(max_rows=200)`、
   `read_plan_cost_rows(max_rows=50)`：上限只作防跑飞，正常靠遇空行退出，撞上限须
   置 `result.warning` 告警而非静默少算。新增循环优先写成 `for`。
2. **可被用户中止** — 外层批量循环每轮开头检查 `_is_sap_cancel_requested()`，在**订单边界**
   `break`（SAP 侧不留半成品单）。`closeEvent` 检测到任务运行中一律 `event.ignore()` +
   置中止标志，**绝不直接 accept**（窗口销毁后循环仍在跑会崩在 `textBrowser.append`）。
3. **不无限重试/等待** — SAP 弹窗、状态轮询、重连一律设轮数上限，超限即失败返回，
   禁止 `while True` 等 SAP 就绪。
4. **不无限增长** — 跳过/新增判定不能让同一条数据每轮重复新增（如 item 号已存在但
   物料不同必须跳过，否则 item 每轮增长）。
5. **异常不回退到重跑同一行** — per-row `except` 内必须 `continue`，不得 retry；
   并写入 `Update Time`，否则出错行与"根本没轮到的行"在 log 里无法区分。

## SAP 身份判定硬约束（基本原则，不可违反）

**任何"这是哪一行/哪个角色/哪个选项"的判定，一律用 key(ID)，绝不用界面显示文本。**

SAP 的中文翻译**不是单射**——多个不同的 key 可以译成同一个中文词，靠文本根本区分不出来。
已确认的实例：`WE`(Ship-to party) 与 `ZG`(Global Partner) 的中文显示**都是"送达方"**。
按文本认角色的后果是把数据挂到错误的伙伴身上，且**不报错、静默写错**，比直接失败危险得多
（2026-09-10 修复：`_GPC_TEXTS` 曾把 WE 的译名"送达方"当作 ZG 的白名单）。

写下拉框/表格行判定前逐条自检：

1. **读现值用 key** — 下拉框对比一律 `find(id).key` 或 `read_key()`，不用 `read_text()`。
   `text` 是显示文本（如 `"E1 国内电商"`），与写入用的 key（`"E1"`）不同口径，
   用 text 对比会永远判定"有差异"而每轮重写。
2. **定位行用 key 扫描** — 表格中找某个角色/类别的行，扫 key 列而非文本列。
   参考 `_find_partner_row(prefix, "ZG")`：命中即证明角色正确，连 `set_key` 都不必调。
   物理行号不能当身份（详见 Item/Plan Cost 的同类约束）。
3. **必须从文本反推 key 时，唯一命中才采信** — 走 `SapSession.list_combo_entries()`
   枚举选项表，命中 0 项或 ≥2 项一律弃权返回 `None`（参考 `_lookup_role_key`）。
   **歧义时宁可记「待校正」不写，也不能猜一个。**
4. **key 写入失败时不得用文本兜底放行** — `set_key` 被拒说明角色没改成，此时回读显示文本
   命中白名单就继续写值，正是上面那个 bug 的成因。一律不写值，保持 SAP 侧原样。
5. **写后回读自证也用 key** — `set_key` 不抛异常 ≠ 生效（SAP 可能悄悄改回），
   回读 key 不等于期望值即视为失败。
6. **显示文本只允许出现在两处** — ① 反查 key 的匹配集；② 选项表不可用（老版 SAP GUI
   无 `Entries`）时的兜底扫描。且兜底仅限**译名无歧义**的角色（如"负责雇员"）。

## 项目说明

### 项目目标
该项目旨在提供一个自动操作SAP创建订单的解决方案，同时还能够获取并处理数据，并实现PDF文件重命名的功能。

### 功能特点
* 数据处理：根据特殊开票信息将原始数据分开，并按要求合并数据。
* SAP操作：实现SAP自动创建订单的功能。
* 数据恢复：能够找回数据，确保数据完整。

### 安装指南
1. 克隆项目：
`git clone https://github.com/chen-huai/Sap_Operation_HDL.git`
2. 安装依赖库：
`pip install -r requirements.txt`
3. 运行主程序：
直接运行主程序`Sap_Operate_HDL.py`即可。
4. 配置文件：
运行成功后，将在桌面生成一个`config`文件夹，并生成`config_sap.csv`的配置文件。可根据`config_sap.csv`文件设置自己需要的参数。
5. 打包成exe程序：
   * 安装第三方库pyinstaller：
   `pip install pyinstaller`
   * 执行打包命令：
   `pyinstaller -F -w 主程序绝对路径`
   
### 源码结构说明

**主程序与 UI**
* `Sap_Operate_HDL.py`：主程序入口，`MyMainWindow` 多 mixin 继承，负责事件绑定与应用启动。
* `Sap_Operate_Ui.py`：Qt Designer 生成的 UI 代码（请勿手动编辑）。
* `Data_Table.py` / `Table_Ui.py`：表格窗口与对应 UI。
* `theme_manager_theme.py`：主题管理。

**业务 Mixin 层**
* `main_window_ui_mixin.py`：UI 初始化、主题、版本菜单。
* `config_mixin.py`：配置文件加载、`getGuiData` 读取 GUI 字段。
* `sap_order_mixin.py`：`odmDataToSap` 多 sheet Excel 批量创建订单、`orderUnlockOrLock` 订单锁/解锁。
* `odm_invoice_mixin.py`：ODM 数据合并/拆分、PDF 发票与电子发票重命名。
* `hour_mixin.py`：`hourOperate` 工时数据批量录入 SAP。

**SAP 服务包 `sap/`**
* `sap/session.py`：`SapSession` 通过 win32com 连接 SAP GUI。
* `sap/services/`：`OrderService`、`InvoiceService`、`HourService` —— 业务级 API。
* `sap/transactions/`：`OrderTransaction`、`InvoiceTransaction`、`HourTransaction` —— 具体 SAP 屏幕操作。
* `sap/models.py`：`SapConfig`、`OrderData`、`RevenueData`、`SapResult` 等 dataclass。
* `sap/rules.py`：A2 物料拆分、Data A 判定、Plan Cost 阈值等业务规则。

**数据 / PDF / 营收**
* `Get_Data.py`：Excel/CSV 数据读取与多 sheet 取数。
* `Excel_Field_Mapper.py`：多命名风格字段统一映射。
* `PDF_Parser_Utils.py`：PDF 读取与发票字段提取（公司名、金额、发票号）。
* `Revenue_Operate.py`：收入分配和工时分配，实现部门收入计算与工时分配。

**基础设施**
* `Logger.py`：基于 pandas 的操作日志记录。
* `chicon.py`：内嵌图标 base64 数据。
* `runtime_globals.py`：跨模块共享全局变量。
* `auto_updater/`：基于 GitHub Releases 的自动更新系统。

## Revenue_Operate.py 收入分配和工时分配逻辑详解

### 模块概述
`Revenue_Operate.py` 是SAP自动化系统中的核心收入分配模块，负责将订单收入自动分配到不同部门和实验室，并计算相应的工时。

### 核心类：RevenueAllocator

#### 1. 基础计算参数设置

**计划成本参数**：
```python
Plan_Cost_Parameter = 0.9  # 实际收入的90%，预留10%利润
```

**有效位数设置**：
```python
Significant_Digits = 2  # 保留2位有效数字
```

**业务部门设置**：
```python
Business_Department = "CS"  # 默认业务部门
```

#### 2. 实验室参数配置

**实验室类型映射**：
```python
Lab_1 = "PHY"      # 物理实验室
Lab_2 = "CHM"      # 化学实验室
T20 = "PHY"        # T20项目对应物理实验室
T75 = "CHM"        # T75项目对应化学实验室
```

**实验室成本比例**：
```python
CHM_Cost_Parameter = 0.3  # 化学实验室获得30%的收入
PHY_Cost_Parameter = 0.3  # 物理实验室获得30%的收入
```

**实验室工时率（每年更新）**：
```python
CS_Hourly_Rate = 329.14    # CS部门每小时费率
CHM_Hourly_Rate = 342      # 化学实验室每小时费率
PHY_Hourly_Rate = 334.75   # 物理实验室每小时费率
```

#### 3. 收入分配计算逻辑

**基础收入计算公式**：
```python
base = (float(revenueData['Revenue']) - float(revenueData['Total Subcon Cost']) / 1.06) * float(configContent.get('Plan_Cost_Parameter'))
```

**分配策略**：

**情况1：物料编码无特殊配置**
- 使用前缀确定实验室类型（T20→PHY，T75→CHM）
- 100%分配给单一实验室和业务部门
- 业务部门获得70%，实验室获得30%

**情况2：物料编码有特殊配置**
- 按项目类型分配比例（405: 50%/50%，441: 80%/20%，430: 80%/20%）
- 1000和2000项目分别对应不同实验室
- 支持交叉分配（如PHY_1000/CHM_2000）

#### 4. 物料编码特殊配置规则

**项目类型分配比例**：
```python
# 405项目分配比例
405_Item_1000 = 0.5  # 1000项目占50%
405_Item_2000 = 0.5  # 2000项目占50%

# 441项目分配比例  
441_Item_1000 = 0.8  # 1000项目占80%
441_Item_2000 = 0.2  # 2000项目占20%

# 430项目分配比例
430_Item_1000 = 0.8  # 1000项目占80%
430_Item_2000 = 0.2  # 2000项目占20%
```

**物料编码实验室映射**：
```python
# T20-430-A2 配置
T20-430-A2 = "PHY_1000/CHM_2000"        # 1000项目分配给PHY，2000项目分配给CHM
T20-430-A2_mc = "T20-430-00/T75-430-00"  # 对应的物料编码

# T75-441-A2 配置
T75-441-A2 = "CHM_1000/PHY_2000"        # 1000项目分配给CHM，2000项目分配给PHY
T75-441-A2_mc = "T75-441-00/T20-441-00"  # 对应的物料编码

# T75-405-A2 配置
T75-405-A2 = "CHM_1000/PHY_2000"        # 1000项目分配给CHM，2000项目分配给PHY
T75-405-A2_mc = "T75-405-00/T20-405-00"  # 对应的物料编码
```

#### 5. 成本中心和工时设置

**成本中心配置**：
```python
CS_Cost_Center = "48601240"    # CS部门成本中心
CHM_Cost_Center = "48601293"   # 化学实验室成本中心
PHY_Cost_Center = "48601294"   # 物理实验室成本中心
```

**工时限制设置**：
```python
Max_Hour = 8  # 每人每天最大工作时长
```

#### 6. 部门数据输出格式

每个订单最终生成4条记录：

1. **1000项目业务部门**：包含订单号、物料编码、业务部门收入和工时
2. **1000项目实验室**：包含订单号、物料编码、实验室收入和工时
3. **2000项目业务部门**：包含订单号、物料编码、业务部门收入和工时
4. **2000项目实验室**：包含订单号、物料编码、实验室收入和工时

#### 7. 计算示例

假设订单数据：
```python
revenueData = {
    'Order Number': 'ORD123',
    'Material Code': 'T20-430-A2',
    'Revenue': 10000,
    'Total Subcon Cost': 1000,
    'Primary CS': 'Chen, Iris'
}
```

**计算过程**：
1. 基础收入 = (10000 - 1000/1.06) × 0.9 = 8150.94元
2. 按项目类型分配（430项目比例：1000项目80%，2000项目20%）：
   - 1000项目收入 = 8150.94 × 0.8 = 6520.75元
   - 2000项目收入 = 8150.94 × 0.2 = 1630.19元
3. 按实验室分配（T20-430-A2配置：PHY_1000/CHM_2000）：
   - **1000项目分配给PHY实验室**：
     - PHY实验室收入 = 6520.75 × 0.3 = 1956.23元
     - CS业务部门收入 = 6520.75 × 0.7 = 4564.52元
   - **2000项目分配给CHM实验室**：
     - CHM实验室收入 = 1630.19 × 0.3 = 489.06元
     - CS业务部门收入 = 1630.19 × 0.7 = 1141.13元
4. **最终三个部门的总收入**：
   - CS部门总收入 = 4564.52 + 1141.13 = 5705.65元
   - PHY实验室收入 = 1956.23元
   - CHM实验室收入 = 489.06元
5. **工时计算**：
   - PHY工时 = 1956.23 / 334.75 = 5.84小时
   - CHM工时 = 489.06 / 342 = 1.43小时
   - CS总工时 = 5705.65 / 329.14 = 17.34小时

#### 8. 主要功能方法

- `allocate_department_hours()`：部门收入分配计算
- `allocate_person_hours()`：工时分配给具体员工
- `allocate_person_average_hours()`：按人员平均分配工时
- `generate_work_days()`：生成有效工作日历
- `_load_hours_data()`：加载工时数据
- `_save_hours_data()`：保存工时数据

#### 9. 配置文件驱动

所有计算参数通过 `config_sap.csv` 配置文件管理，包括：
- 基础参数（计划成本参数、有效位数等）
- 实验室参数（成本比例、工时率等）
- 物料编码特殊配置
- 成本中心设置
- 人员信息映射

这种配置驱动的设计使得系统具有高度的灵活性和可维护性，可以根据业务需求快速调整分配规则。

### 特别感谢
* 特别感谢JetBrains的支持，提供优秀的开发工具，让我们能够更高效地进行编码工作。

## Project Description

### Project Objective
This project aims to provide a solution for automatically operating SAP to create orders, as well as obtaining and processing data, and implementing PDF file renaming functionality.

### Features
* Processing: Separate original data based on special invoicing information and merge data as required.
* SAP Operations: Implement the functionality to automatically create orders in SAP.
* Data Recovery: Ability to retrieve data to ensure data integrity.

### Installation Guide
1. Clone the project:
`git clone https://github.com/chen-huai/Sap_Operation_HDL.git`
2. Install dependencies:
`pip install -r requirements.txt`
3. Run the main program: Simply run the main program Sap_Operate_HDL.py.
4. Configuration file: Upon successful execution, a `config` folder will be generated on the desktop, along with a `config_sap.csv` configuration file. You can set your own parameters in the `config_sap.csv` file.
5. Package as an exe program:
   * Install the third-party library pyinstaller:
   `pip install pyinstaller`
   * Execute the packaging command:
   `pyinstaller -F -w absolute_path_to_main_program`
   
### Source Code Structure

**Main program & UI**
* `Sap_Operate_HDL.py`: Main entry — `MyMainWindow` with multi-mixin inheritance, wiring events and bootstrapping the app.
* `Sap_Operate_Ui.py`: UI generated by Qt Designer (do not hand-edit).
* `Data_Table.py` / `Table_Ui.py`: Table window and its UI.
* `theme_manager_theme.py`: Theme management.

**Business mixins**
* `main_window_ui_mixin.py`: UI initialization, themes, version menu.
* `config_mixin.py`: Config file loading, `getGuiData` for GUI field readout.
* `sap_order_mixin.py`: `odmDataToSap` multi-sheet Excel batch order creation; `orderUnlockOrLock` for order lock/unlock.
* `odm_invoice_mixin.py`: ODM data combine/split, PDF invoice and electronic invoice renaming.
* `hour_mixin.py`: `hourOperate` batch hour-entry into SAP.

**SAP service package `sap/`**
* `sap/session.py`: `SapSession` win32com connection wrapper.
* `sap/services/`: `OrderService`, `InvoiceService`, `HourService` — business-level APIs.
* `sap/transactions/`: `OrderTransaction`, `InvoiceTransaction`, `HourTransaction` — concrete SAP screen automation.
* `sap/models.py`: `SapConfig`, `OrderData`, `RevenueData`, `SapResult` and other dataclasses.
* `sap/rules.py`: A2 material split, Data A customer matching, Plan Cost thresholds.

**Data / PDF / Revenue**
* `Get_Data.py`: Excel/CSV data reading and multi-sheet access.
* `Excel_Field_Mapper.py`: Cross-style field-name normalization.
* `PDF_Parser_Utils.py`: PDF reading and invoice field extraction (company, amount, invoice no.).
* `Revenue_Operate.py`: Revenue and hour allocation engine.

**Infrastructure**
* `Logger.py`: pandas-backed operation log writer.
* `chicon.py`: Embedded icon base64 data.
* `runtime_globals.py`: Cross-module shared globals.
* `auto_updater/`: GitHub-Releases-driven auto update system.

### Special Thanks
* Special thanks to JetBrains for their support, providing excellent development tools that allow us to code more efficiently.
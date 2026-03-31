## 项目说明

### 项目目标
该项目旨在提供一个自动操作SAP创建订单的解决方案，同时还能够获取并处理数据，并实现PDF文件重命名的功能。

### 功能特点
* 数据处理：根据特殊开票信息将原始数据分开，并按要求合并数据。
* SAP操作：实现SAP自动创建订单的功能。
* 数据恢复：能够找回数据，确保数据完整。

### 安装指南
1. 克隆项目：
`git clone https://github.com/chen-huai/Sap_Operation.git`
2. 安装依赖库：
`pip install -r requirements.txt`
3. 运行主程序：
直接运行主程序`Sap_Operate.py`即可。
4. 配置文件：
运行成功后，将在桌面生成一个`config`文件夹，并生成`config_sap.csv`的配置文件。可根据`config_sap.csv`文件设置自己需要的参数。
5. 打包成exe程序：
   * 安装第三方库pyinstaller：
   `pip install pyinstaller`
   * 执行打包命令：
   `pyinstaller -F -w 主程序绝对路径`
   
### 源码结构说明
* `Sap_Operate.py`：主程序，包含基础配置、数据处理、SAP逻辑操作、PDF重命名的逻辑操作。
* `Sap_Operate_Ui.py`：UI界面。
* `Get_Data.py`：数据处理模块，包括数据基础处理。
* `Sap_Function.py`：SAP操作模块，实现SAP基础操作。
* `File_Operate.py`：文件处理模块，用于创建文件夹和获取文件名称。
* `Data_Table.py`：表格处理模块，实现表格基础设置。
* `PDF_Operate.py`：PDF文件处理模块，包含PDF文件读取和保存功能。
* `Revenue_Operate.py`：收入分配和工时分配模块，实现自动化的部门收入计算和工时分配。

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
`git clone https://github.com/chen-huai/Sap_Operation.git`
2. Install dependencies:
`pip install -r requirements.txt`
3. Run the main program: Simply run the main program Sap_Operate.py.
4. Configuration file: Upon successful execution, a `config` folder will be generated on the desktop, along with a `config_sap.csv` configuration file. You can set your own parameters in the `config_sap.csv` file.
5. Package as an exe program:
   * Install the third-party library pyinstaller:
   `pip install pyinstaller`
   * Execute the packaging command:
   `pyinstaller -F -w absolute_path_to_main_program`
   
### Source Code Structure
* `Sap_Operate.py`: Main program, including basic configuration, data processing, SAP logic operations, and PDF renaming logic operations.
* `Sap_Operate_Ui.py`: UI interface.
* `Get_Data.py`: Data processing module, including basic data processing.
* `Sap_Function.py`: SAP operation module, implementing basic SAP operations.
* `File_Operate.py`: File processing module for creating folders and obtaining file names.
* `Data_Table.py`: Table processing module, implementing basic table settings.
* `PDF_Operate.py`: PDF file processing module, including PDF file reading and saving functionality.

### Special Thanks
* Special thanks to JetBrains for their support, providing excellent development tools that allow us to code more efficiently.
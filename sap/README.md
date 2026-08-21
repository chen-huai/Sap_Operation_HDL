# SAP 模块说明

## 模块用途

`sap` 模块是一个面向业务流程的 SAP GUI 自动化模块。

它的目标不是让调用方直接操作 `findById(...)`，而是把 SAP 自动化拆成清晰的几层，让上层代码按下面的方式使用：

1. 建立 SAP 会话
2. 选择业务域服务
3. 传入业务数据
4. 显式执行步骤

当前模块主要覆盖四类业务：

- 订单创建业务：创建订单、打开订单、添加 item、填写 Data B 人工成本、填写计划成本、保存、锁定、解锁
- 订单编辑业务（VA02 对比更新）：读 SAP 现值与 Excel 对比，**只改差异**——抬头、item、计划成本、Data B、订单价值
- 发票业务：创建形式发票、查看形式发票
- 工时业务：登录工时系统、录入工时、保存工时

## 为什么要这样拆

旧结构的问题是：

- 业务规则和 SAP 页面操作混在一起
- 一个类同时承担连接、判断、流程编排、页面点击
- 调用方必须知道很多隐式前置状态

新结构的设计目标是：

- 让上层看到的是“订单服务”“工时服务”“发票服务”
- 让业务规则单独放在规则层
- 让 SAP 页面细节留在事务层
- 让主程序不再依赖一个庞大的单体类

## 如何理解这个模块

如果只从“使用者”的角度理解，这个模块可以简化成一句话：

“这是一个把 SAP GUI 自动化包装成服务接口的业务模块。”

你只需要理解三件事：

1. `SapSession`
作用：代表当前 SAP GUI 会话

2. `Service`
作用：代表你想做的业务域

- `OrderService`（创建域）
- `OrderEditService`（编辑域，VA02 对比更新）
- `InvoiceService`
- `HourService`

3. `models`
作用：代表传给模块的业务数据和选项

## 模块分层

### 1. 会话层

文件：[session.py](/C:/Data/Python/Sap_Operation_HDL/sap/session.py)

用途：

- 连接 SAP GUI
- 获取 COM session
- 统一封装查找控件、读写文本、点击按钮、切换页签、发送快捷键

你可以把这一层理解成“驱动层”。

### 2. 规则层

文件：[rules.py](/C:/Data/Python/Sap_Operation_HDL/sap/rules.py)

用途：

- 判断物料类型
- 解析 A2 物料映射
- 决定 DATA A 的 key
- 计算 lab cost / plan cost 应该生成哪些行
- 判断是否需要写订单价值、是否需要计划成本

这一层只做“业务判断”，不做页面点击。

你可以把这一层理解成“业务脑子”。

### 3. 事务层

目录：[transactions](/C:/Data/Python/Sap_Operation_HDL/sap/transactions)

用途：

- 把一段具体 SAP 页面操作封装成事务
- 知道页签、字段 ID、按钮、菜单位置

主要文件：

- [order.py](/C:/Data/Python/Sap_Operation_HDL/sap/transactions/order.py)：订单创建相关事务
- [order_edit.py](/C:/Data/Python/Sap_Operation_HDL/sap/transactions/order_edit.py)：订单编辑事务（VA02 读现值-对比-仅改差异）
- [invoice.py](/C:/Data/Python/Sap_Operation_HDL/sap/transactions/invoice.py)：发票相关事务
- [hours.py](/C:/Data/Python/Sap_Operation_HDL/sap/transactions/hours.py)：工时相关事务

`order_edit.py` 通过组合持有一个 `OrderTransaction`，复用其 `open()`/`save()` 与控件 ID helper，
保证「对比口径」与「创建口径」逐字段一致——否则未变化的数据会因两套逻辑差异被误判为有变化而错误重写。

你可以把这一层理解成“会操作 SAP 的执行层”。

### 4. 服务层

目录：[services](/C:/Data/Python/Sap_Operation_HDL/sap/services)

用途：

- 提供给主程序直接调用的明确入口
- 按业务域组织能力，而不是按 SAP 事务码暴露细节

主要文件：

- [order_service.py](/C:/Data/Python/Sap_Operation_HDL/sap/services/order_service.py)：订单创建域
- [order_edit_service.py](/C:/Data/Python/Sap_Operation_HDL/sap/services/order_edit_service.py)：订单编辑域（VA02）
- [invoice_service.py](/C:/Data/Python/Sap_Operation_HDL/sap/services/invoice_service.py)
- [hour_service.py](/C:/Data/Python/Sap_Operation_HDL/sap/services/hour_service.py)

你可以把这一层理解成“给业务代码调用的 API 层”。

### 5. 数据模型层

文件：[models.py](/C:/Data/Python/Sap_Operation_HDL/sap/models.py)

用途：

- 定义配置
- 定义订单/营收/工时数据
- 定义操作选项
- 定义统一返回结果

## 公开入口

模块对外只暴露这些入口：

```python
from sap import (
    CostOptions,
    DataBEntry,
    HourData,
    HourService,
    InvoiceService,
    ItemAddInfo,
    OrderData,
    OrderEditService,
    OrderItemData,
    OrderService,
    PartnerOptions,
    PlanCostEntry,
    RevenueData,
    SapConfig,
    SapResult,
    SapSession,
)
```

入口定义见：[__init__.py](/C:/Data/Python/Sap_Operation_HDL/sap/__init__.py)

## 适用场景

### 适合使用这个模块的场景

- 你已经登录了 SAP GUI，需要在现有会话上执行自动化操作
- 你要做的是订单、发票或工时这三类已封装业务
- 你希望主程序调用的是业务动作，而不是底层控件 ID
- 你希望业务规则和 UI 自动化代码分离
- 你需要在后续继续扩展订单流程，但不想再把逻辑堆进一个大类里

### 最适合的具体场景

- 从 Excel、数据库或业务系统读取订单数据后，批量创建 SAP 订单
- 给已存在订单追加 item、补成本、补计划成本
- 对已存在订单做差异化更新：Excel 为准，只改变化的字段，未变化的一律不碰
- 为订单创建或查看形式发票
- 把工时记录批量写入 SAP 工时系统

### 不适合使用这个模块的场景

- 你想做完全不同的 SAP 事务，但当前事务层里还没有封装
- 你想直接操控任意 SAP 页面控件，而不走现有业务流程
- 你运行环境里没有 SAP GUI 或没有启用脚本控制
- 你依赖旧的单体 `Sap` 类接口

## 使用方式

### 第一步：连接 SAP

```python
from sap import SapSession

session = SapSession.connect()
```

### 第二步：选择服务

订单业务：

```python
from sap import OrderService

order_service = OrderService(session, config)
```

发票业务：

```python
from sap import InvoiceService

invoice_service = InvoiceService(session)
```

工时业务：

```python
from sap import HourService

hour_service = HourService(session)
```

### 第三步：传入业务数据

订单类业务会用到：

- `SapConfig`
- `OrderData`
- `RevenueData`
- `PartnerOptions`
- `CostOptions`

工时业务会用到：

- `HourData`

### 第四步：显式调用步骤

新模块不做隐式联动。

这意味着：

- 不会因为你调用了一个方法，就自动顺带调用后续所有步骤
- 每一步都由主程序显式决定

这样做的好处是流程更清楚，也更容易适配新的业务逻辑。

## 典型调用模式

### 订单流程

```python
from sap import (
    CostOptions,
    OrderData,
    OrderService,
    PartnerOptions,
    RevenueData,
    SapConfig,
    SapSession,
)

session = SapSession.connect()

config = SapConfig(
    order_type="ZOR",
    sales_organization="3002",
    distribution_channels="10",
    sales_office="1000",
    cost_center="1100",
    sub_cost_center_cs="1101",
    sub_cost_center_chm="1102",
    sub_cost_center_phy="1103",
    cs_code="CS001",
    sales_code="SA001",
)

service = OrderService(session, config)

order = OrderData(
    sap_no="123456",
    project_no="PRJ-001",
    material_code="T75-405-00",
    currency_type="CNY",
    exchange_rate=1.0,
    cost=5000.0,
    short_text="Short text",
    long_text="Long text",
    global_partner_code="GP001",
    sales_name="Alice",
    ecd="2026.12.31",
)

revenue = RevenueData(
    revenue=10000.0,
    revenue_cny=72500.0,
    chm_cost=3000.0,
    phy_cost=2000.0,
    chm_revenue=5000.0,
    phy_revenue=3000.0,
)

create_result = service.create_order(
    order,
    revenue,
    partner_options=PartnerOptions(add_contact=True, add_sales_partner=True),
)

if create_result.success:
    service.open_order("60001234")
    service.add_items(order, revenue)
    # Data B / 计划成本传「已算好的明细」，规则计算在调用方（或 sap/rules.py）完成
    service.fill_lab_cost_entries(data_b_entries, order)
    service.apply_plan_cost_entries(plan_cost_entries, focus_row=0)
    service.fill_order_value(order)          # Σ SAP item 净值 × 汇率，达阈值才写
    service.save("订单")

session.close()
```

> 注意：`fill_lab_cost_entries()` / `apply_plan_cost_entries()` 接收的是**已计算好的明细列表**
> （`list[DataBEntry]` / `list[PlanCostEntry]`），不是 `RevenueData`。
> 明细由调用方按业务规则先算好，事务层只负责写入——这样规则与页面操作彻底解耦。

### 订单编辑流程（VA02 对比更新）

编辑域与创建域业务分割：编辑只负责「读 SAP 现值 → 与 Excel 对比 → 仅改差异」，绝不无条件重写。

```python
from sap import OrderEditService, SapSession

session = SapSession.connect()
service = OrderEditService(session, config)

service.open_order("60001234")

diffs = []                                    # 差异摘要收集器，各段共享同一套格式
service.edit_header(order, diffs)             # 抬头：售达方/币种/汇率/伙伴/DATA A/ECD/IC
service.edit_items(order, diffs, added_out=added)   # item：双键匹配，仅改金额/币种/长文本

# item 有新增 → SAP 可能自动改号，须建立 ODM→SAP 号映射供下游定位
item_no_map = service.build_item_no_mapping(added)

service.edit_plan_cost(entries, diffs, target_item="1000")   # 计划成本：主键匹配、顺序无关

# Data B 必须两段式，中间隔一次 save + open_order（SAP 硬约束，否则必报 ZR520）
clear = service.clear_data_b(data_b_entries, order, diffs, item_no_map=item_no_map)
if clear.changed:
    service.save("Before Data B")
    service.open_order("60001234")
    item_no_map = service.build_item_no_mapping(added)
    service.write_data_b(data_b_entries, order, diffs, item_no_map=item_no_map)

service.edit_order_value(order, diffs)        # 订单价值 = Σ SAP item 净值 × 汇率
service.save("VA02 Edit")

session.close()
```

#### 对比口径约定（正确性关键）

SAP 回读格式与 Excel 不同，直接文本比对会让「其实没变」的数据每次都被判为差异并重写。
故所有对比统一走归一化原语：

| 原语 | 用途 | 解决的坑 |
|------|------|----------|
| `_norm` | 文本 | 多行文本回读用 `\r\n`，Excel 是 `\n` |
| `_norm_amount` | 金额（.2f） | 回读带千分位 `1,000.00` |
| `_norm_number` | 汇率等数值 | 回读带尾随零 `7.10000` vs `7.1` |
| `_norm_no` | 成本中心 / item 号 | SAP 定长字段回读带前导零 `0048601258` |

**计划成本按 `(成本中心, 类别)` 主键匹配**，对行序免疫：SAP 与 Excel 顺序不同但内容一致时
判为无差异，零写入零回车；金额不同才改 MENGE 列；Excel 有 SAP 无则追加到末尾；SAP 有 Excel 无则删除。
比对天然限定在单个 item 内（编辑器只含该 item 的行），不跨 item。

#### 日志格式约定

各段差异摘要统一格式，便于人工核对：

- 箭头一律 `旧→新`（无空格）：`汇率:7.10000→7.1`、`成本中心48601240(FREMDL) 金额 80.00→100.00`
- 空值一律渲染 `(空)`：`订单价值:3,021.26→(空)` 一眼看出「被清空」，而非结尾空白
- **无差异的段不写入 log**，只保留真实变更、警告与失败；步骤执行情况仍由 UI 步骤面板完整呈现

### 发票流程

```python
from sap import InvoiceService, SapSession

session = SapSession.connect()
service = InvoiceService(session)

service.create_proforma()
service.display_proforma()

session.close()
```

### 工时流程

```python
from sap import HourData, HourService, SapSession

session = SapSession.connect()
service = HourService(session)

hour = HourData(
    staff_id="EMP001",
    week="15",
    allocated_day="2026.04.07",
    order_no="ORD-001",
    item="10",
    material_code="T75-405-00",
    allocated_hours=8.0,
    office_time=8.0,
)

service.login(hour)
service.record(hour, row_num=0, max_rows=20)
service.save(max_retries=14)

session.close()
```

## 常用服务及用途

### OrderService

适用场景：

- 创建订单头
- 打开已有订单
- 添加 item
- 填写 Data B 人工成本
- 写计划成本
- 保存订单
- 锁定/解锁订单

常用方法：

- `create_order()`
- `open_order()`
- `add_items()` / `update_items()`
- `fill_lab_cost_entries(entries, order)`
- `apply_plan_cost_entries(entries, focus_row=0)`
- `fill_order_value(order)`
- `save()`
- `lock()`
- `unlock()`

### OrderEditService

适用场景（VA02，只改差异）：

- 打开已有订单并对比更新抬头字段
- 按 item + 物料双键匹配更新 item（金额/币种/长文本），SAP 有 Excel 无的行只提示不删改
- 按主键匹配更新计划成本（顺序无关）
- 两段式同步 Data B（删空与重建之间必须隔一次保存）
- 重算回填订单价值

常用方法：

- `open_order()`
- `edit_header()`
- `edit_items()`
- `build_item_no_mapping()`：新增 item 被 SAP 自动改号后建立 ODM→SAP 号映射
- `edit_plan_cost()`
- `clear_data_b()` / `write_data_b()`：**必须成对使用且中间隔一次 `save()` + `open_order()`**
- `edit_order_value()`
- `save()`

### InvoiceService

适用场景：

- 创建形式发票
- 查看形式发票

常用方法：

- `create_proforma()`
- `display_proforma()`

### HourService

适用场景：

- 登录工时系统
- 录入工时
- 保存工时

常用方法：

- `login()`
- `record()`
- `save()`

## 返回结果怎么理解

所有服务方法基本都返回 `SapResult`。

重点字段：

- `success`：是否成功
- `message`：失败或补充信息
- `step`：当前结果对应的步骤名
- `order_no`：订单号
- `proforma_no`：形式发票号
- `sap_amount_vat`：SAP 页面返回的金额文本
- `warning`：未失败但需提示用户（如「SAP 无对应 item，已跳过」），UI 用区别色显示
- `changed`：本步骤是否**实际改动了** SAP；`False` 表示对比后无差异、原样跳过。
  编辑流程据此决定后续动作（如 `clear_data_b()` 返回 `changed=False` 则免掉中途保存）

注意：`changed` 表达的是「是否动了 SAP」，不要用它判断「是否有内容值得写日志」——
例如「SAP 有、Excel 无，已跳过」没改 SAP 但必须留痕，后者应看该段的差异摘要是否为空。

推荐用法：

```python
result = service.create_order(order, revenue)
if not result.success:
    print(result.step)
    print(result.message)
```

## 和旧接口的区别

新模块不再提供单体式 `Sap` 类，也不再支持旧的 `sap/function.py` 使用方式。

旧入口 [Sap_Function.py](/C:/Data/Python/Sap_Operation_HDL/Sap_Function.py) 已明确废弃，并会直接提示改用新入口。

## 一句话总结

如果你是调用方，可以把这个模块理解成：

“一个把 SAP GUI 自动化封装成订单创建、订单编辑、发票、工时四类服务接口的业务模块。”

# SAP 快速接入

## 一句话理解

`sap` 模块是一个把 SAP GUI 自动化封装成业务服务的模块。

你不再直接操作 SAP 控件，而是按下面的模式使用：

1. 建立 `SapSession`
2. 选择业务服务
3. 传入数据
4. 显式调用步骤

## 公开入口

```python
from sap import (
    SapSession,
    SapConfig,
    OrderData,
    OrderItemData,
    RevenueData,
    HourData,
    DataBEntry,
    PlanCostEntry,
    ItemAddInfo,
    PartnerOptions,
    CostOptions,
    OrderService,
    OrderEditService,
    InvoiceService,
    HourService,
)
```

## 四个服务分别做什么

### `OrderService`

适用场景（创建域，VA01/VA02 写入）：

- 创建订单
- 打开订单
- 添加 item
- 填 Data B 人工成本
- 填计划成本
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

### `OrderEditService`

适用场景（编辑域，VA02 对比更新——**只改差异**）：

- 对比更新抬头字段
- 对比更新 item（item + 物料双键匹配）
- 对比更新计划成本（`(成本中心, 类别)` 主键匹配，顺序不同不算变化）
- 两段式同步 Data B
- 重算回填订单价值

常用方法：

- `open_order()`
- `edit_header()`
- `edit_items()`
- `build_item_no_mapping()`
- `edit_plan_cost()`
- `clear_data_b()` / `write_data_b()`
- `edit_order_value()`
- `save()`

### `InvoiceService`

适用场景：

- 创建形式发票
- 查看形式发票

常用方法：

- `create_proforma()`
- `display_proforma()`

### `HourService`

适用场景：

- 登录工时系统
- 录入工时
- 保存工时

常用方法：

- `login()`
- `record()`
- `save()`

## 最常见的调用模式

### 订单

```python
from sap import (
    SapSession,
    SapConfig,
    OrderData,
    RevenueData,
    PartnerOptions,
    OrderService,
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
)

revenue = RevenueData(
    revenue=10000.0,
    revenue_cny=72500.0,
)

result = service.create_order(
    order,
    revenue,
    partner_options=PartnerOptions(),
)

if result.success:
    service.open_order("60001234")
    service.add_items(order, revenue)
    # 传已算好的明细列表，不是 RevenueData
    service.fill_lab_cost_entries(data_b_entries, order)
    service.apply_plan_cost_entries(plan_cost_entries, focus_row=0)
    service.fill_order_value(order)
    service.save("订单")

session.close()
```

### 订单编辑（只改差异）

```python
from sap import SapSession, OrderEditService

session = SapSession.connect()
service = OrderEditService(session, config)

service.open_order("60001234")

diffs = []                              # 差异摘要收集器；为空表示该段无变化
service.edit_header(order, diffs)
service.edit_items(order, diffs)
service.edit_plan_cost(plan_cost_entries, diffs, target_item="1000")

# Data B 两段式：删空与重建之间必须隔一次 save + open_order，否则 SAP 必报 ZR520
clear = service.clear_data_b(data_b_entries, order, diffs)
if clear.changed:
    service.save("Before Data B")
    service.open_order("60001234")
    service.write_data_b(data_b_entries, order, diffs)

service.edit_order_value(order, diffs)
service.save("VA02 Edit")

session.close()
```

### 发票

```python
from sap import SapSession, InvoiceService

session = SapSession.connect()
service = InvoiceService(session)

service.create_proforma()
service.display_proforma()

session.close()
```

### 工时

```python
from sap import SapSession, HourData, HourService

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
service.record(hour)
service.save()

session.close()
```

## 记住这几点

- 新模块没有旧的单体 `Sap` 类。
- 新模块不会隐式帮你串完整流程，后续步骤要显式调用。
- 大多数调用都会返回 `SapResult`，先看 `success`，失败时看 `message` 和 `step`；
  `warning=True` 表示未失败但需提示（如「SAP 无对应 item，已跳过」），`changed=False` 表示对比后无差异、原样跳过。
- 编辑域的对比一律走归一化（金额去千分位、编号去前导零、汇率按数值），
  否则「其实没变」的数据会因回读格式差异被误判为变化而每次重写。
- 编辑域的差异日志统一 `旧→新` 格式，空值渲染 `(空)`；无差异的段不产生日志。
- 如果你要完整说明，看 [README.md](/C:/Data/Python/Sap_Operation_HDL/sap/README.md)。

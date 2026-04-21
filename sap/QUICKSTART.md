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
    RevenueData,
    HourData,
    PartnerOptions,
    CostOptions,
    OrderService,
    InvoiceService,
    HourService,
)
```

## 三个服务分别做什么

### `OrderService`

适用场景：

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
- `add_items()`
- `fill_lab_cost()`
- `apply_plan_cost()`
- `save()`
- `lock()`
- `unlock()`

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
    CostOptions,
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
    service.fill_lab_cost(order, revenue)
    service.apply_plan_cost(order, revenue, cost_options=CostOptions())
    service.save("订单")

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
- 大多数调用都会返回 `SapResult`，先看 `success`，失败时看 `message` 和 `step`。
- 如果你要完整说明，看 [README.md](/C:/Data/Python/Sap_Operation_HDL/sap/README.md)。

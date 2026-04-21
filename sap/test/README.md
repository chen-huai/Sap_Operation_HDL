# SAP 测试说明

## 测试分层

`sap/test` 已同步到新的模块结构，分成两类测试：

1. 单元测试
用途：验证 `session / rules / transactions / services` 的逻辑  
特点：全部基于 mock，不需要真实 SAP GUI

2. 集成测试
用途：在真实 SAP GUI 上做冒烟验证  
特点：默认跳过，只有显式开启才运行

## 默认运行方式

```bash
pytest sap/test -q
```

这条命令默认只跑单元测试。

## 运行集成测试

```bash
pytest sap/test --run-integration -q
```

如果不加 `--run-integration`，`sap/test/integration` 下的测试会被自动跳过。

## 测试目录

```text
sap/test/
  helpers.py                 # 测试工厂和 mock 工具
  conftest.py                # pytest 配置和 fixtures
  test_rules.py              # 规则层测试
  test_session.py            # 会话层测试
  test_order_service.py      # 订单服务/事务测试
  test_invoice_service.py    # 发票服务测试
  test_hour_service.py       # 工时服务测试
  integration/
    test_config.py           # 集成测试参数
    conftest.py              # 真实 SAP session fixtures
    test_connection.py       # 连接冒烟
    test_order_service.py    # 订单流程冒烟
    test_invoice_service.py  # 发票流程冒烟
    test_hour_service.py     # 工时流程冒烟
```

## helpers.py 常用工厂

- `make_config()`
- `make_order()`
- `make_revenue()`
- `make_hour()`
- `make_partner_options()`
- `make_cost_options()`
- `create_raw_session()`
- `create_sap_session()`
- `create_order_service()`
- `create_invoice_service()`
- `create_hour_service()`

## 编写新测试的建议

- 规则判断优先写在 `test_rules.py`
- Service 层测试优先验证业务入口和结果对象
- 事务层需要验证控件写入时，使用 `create_raw_session()` 读取缓存元素
- 真实 SAP 验证只放到 `integration/`

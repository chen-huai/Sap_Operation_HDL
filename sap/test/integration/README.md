# SAP 集成测试

这些测试会连接真实 SAP GUI，会执行真实页面操作。

## 运行方式

```bash
pytest sap/test --run-integration -q
```

也可以只跑单个集成测试：

```bash
pytest sap/test/integration/test_connection.py --run-integration -q -s
```

## 使用前准备

1. 已安装并登录 SAP GUI
2. 已开启脚本控制
3. 已填写 [test_config.py](/C:/Data/Python/Sap_Operation_HDL/sap/test/integration/test_config.py)
4. 确认测试数据不会影响生产业务

## 当前覆盖范围

- `test_connection.py`：SAP 连接冒烟
- `test_order_service.py`：订单服务冒烟
- `test_invoice_service.py`：发票服务冒烟
- `test_hour_service.py`：工时服务冒烟

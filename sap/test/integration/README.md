# SAP 集成测试

连接真实 SAP GUI 环境，验证每个功能是否正常执行。

## 快速开始

```bash
# 1. 编辑测试参数
#    打开 sap/test/integration/test_config.py，填写你的测试数据

# 2. 确保 SAP GUI 已登录

# 3. 运行单个测试
pytest sap/test/integration/test_connection.py -v -s
```

## 安全机制

| 控制项 | 默认值 | 说明 |
|--------|--------|------|
| `ALLOW_SAVE` | `False` | 保存类测试（save_sap, save_hours）需手动改为 `True` |
| 参数未填写 | 自动 SKIP | 如 ORDER.sap_no 为空，相关测试自动跳过 |

## 测试文件与依赖关系

### 独立运行（无前置依赖）

| 测试文件 | 需要配置 | 说明 |
|----------|----------|------|
| `test_connection.py` | 无 | 仅验证 SAP 连接 |
| `test_open_va02.py` | `EXISTING_ORDER_NO` | 打开已有订单 |
| `test_end_sap.py` | 无 | 创建独立连接再断开 |
| `test_login_hour_gui.py` | `HOUR` | 登录工时系统 |

### 需要前置操作

| 测试文件 | 前置操作 | 需要配置 |
|----------|----------|----------|
| `test_va01_operate.py` | 无 | `ORDER`, `REVENUE` |
| `test_lab_cost.py` | VA01 之后 | `ORDER`, `REVENUE` |
| `test_va02_operate.py` | VA01 之后 或 有 `EXISTING_ORDER_NO` | `ORDER`, `REVENUE` |
| `test_plan_cost.py` | VA02 之后 | `ORDER`, `REVENUE` |
| `test_save_sap.py` | 任意操作之后 | `ALLOW_SAVE = True` |
| `test_vf01_operate.py` | VA02 + save 之后 | 无 |
| `test_vf03_operate.py` | VF01 + save 之后 | 无 |
| `test_unlock_or_lock_order.py` | 无（自动 open） | `EXISTING_ORDER_NO` |
| `test_recording_hours.py` | login_hour 之后 | `HOUR` |
| `test_save_hours.py` | recording 之后 | `ALLOW_SAVE = True` |

## 典型测试流程

### 订单完整流程

```bash
# 按顺序执行（同一 pytest 会话共享 SAP 连接）
pytest sap/test/integration/test_va01_operate.py \
      sap/test/integration/test_lab_cost.py \
      sap/test/integration/test_va02_operate.py \
      sap/test/integration/test_save_sap.py \
      sap/test/integration/test_vf01_operate.py \
      sap/test/integration/test_save_sap.py \
      sap/test/integration/test_vf03_operate.py \
      -v -s
```

### 工时完整流程

```bash
pytest sap/test/integration/test_login_hour_gui.py \
      sap/test/integration/test_recording_hours.py \
      sap/test/integration/test_save_hours.py \
      -v -s
```

### 单独验证某个功能

```bash
# PyCharm: 右击测试文件 → Run
# 命令行:
pytest sap/test/integration/test_open_va02.py -v -s
```

## 失败时的输出

测试失败会显示 SAP 返回的实际错误信息：

```
FAILED test_va01_operate.py::TestVa01Operate::test_va01_create_order
  AssertionError: VA01 失败: Order No未创建成功，找不到元素 wnd[0]/usr/ctxtVBAK-AUART
```

根据错误信息定位 `sap/function.py` 中的具体步骤进行修复。

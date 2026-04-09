# SAP Function 测试模块

## 快速开始

```bash
# 进入项目根目录
cd C:\Data\Python\Sap_Operation_HDL

# 安装 pytest（仅首次）
pip install pytest

# 运行全部测试
pytest sap/test/ -v

# 运行单个测试文件
pytest sap/test/test_va01_operate.py -v

# 运行单个测试类
pytest sap/test/test_va01_operate.py::TestVa01OperateCurrency -v

# 运行单个测试用例
pytest sap/test/test_va01_operate.py::TestVa01OperateCurrency::test_cny_no_exchange_rate -v
```

## 文件结构

```
sap/test/
├── helpers.py                     # 工厂函数和 Mock 工具（公共基础设施）
├── conftest.py                    # pytest fixtures（自动加载）
│
│  ── 订单流程 ──
├── test_init.py                   # Sap.__init__     SAP 连接初始化
├── test_va01_operate.py           # va01_operate      创建订单
├── test_va02_operate.py           # va02_operate      添加/修改 Item
├── test_save_sap.py               # save_sap          保存数据
├── test_open_va02.py              # open_va02         打开订单
├── test_unlock_or_lock_order.py   # unlock_or_lock    解锁/锁定订单
├── test_end_sap.py                # end_sap           结束连接
│
│  ── 成本处理 ──
├── test_lab_cost.py               # lab_cost          Data B 成本填写
├── test_plan_cost.py              # plan_cost         成本计划
│
│  ── 发票 ──
├── test_vf01_operate.py           # vf01_operate      创建形式发票
├── test_vf03_operate.py           # vf03_operate      查看形式发票
│
│  ── 工时管理 ──
├── test_login_hour_gui.py         # login_hour_gui    登录工时系统
├── test_recording_hours.py        # recording_hours   记录工时
├── test_save_hours.py             # save_hours        保存工时
│
│  ── 纯逻辑 ──
└── test_get_a2_materials.py       # _get_a2_materials A2 物料映射
```

## 测试原理

所有测试通过 Mock 模拟 SAP GUI COM 接口，**无需启动 SAP 即可运行**。

核心机制：
- `helpers.create_sap_instance()` — 跳过 `__init__` 的 COM 连接，注入 Mock session
- `helpers.setup_session_mock()` — 配置 `session.findById()` 对不同 ID 返回不同值
- `helpers.make_*()` — 工厂函数，快速创建测试数据

## 常用命令

```bash
# 查看详细输出（含每个 assert）
pytest sap/test/ -v

# 遇到第一个失败立即停止
pytest sap/test/ -x

# 只运行上次失败的测试
pytest sap/test/ --lf

# 显示最慢的 5 个测试
pytest sap/test/ --durations=5

# 按关键字过滤（运行所有含 "a2" 的测试）
pytest sap/test/ -k "a2"

# 按关键字过滤（运行所有异常测试）
pytest sap/test/ -k "exception"
```

## 编写新测试

当 `sap/function.py` 新增方法时，按以下模板创建测试文件：

```python
"""
测试 Sap.新方法名 — 功能说明

使用方法:
    pytest sap/test/test_新方法名.py -v
"""

from unittest.mock import MagicMock

from sap.test.helpers import (
    create_sap_instance, make_order, make_revenue, make_flags,
    setup_session_mock,
)


class TestNewMethod:
    """新方法名 功能说明"""

    def test_success(self):
        """正常路径"""
        session = MagicMock()
        # 如果方法内部需要读取 session 元素的 .text 值：
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': '期望的状态栏文本',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.新方法名(参数)

        assert result.success

    def test_exception(self):
        """异常路径"""
        session = MagicMock()
        session.findById.side_effect = Exception("模拟异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.新方法名(参数)

        assert not result.success
        assert '期望的错误关键字' in result.message
```

### helpers.py 工厂函数速查

| 函数 | 用途 | 示例 |
|------|------|------|
| `make_config(**kw)` | 创建 SapConfig | `make_config(revenue_threshold=50000)` |
| `make_order(**kw)` | 创建 OrderData | `make_order(material_code='D2-405-00')` |
| `make_revenue(**kw)` | 创建 RevenueData | `make_revenue(revenue_cny=20000.0)` |
| `make_flags(**kw)` | 创建 OperationFlags | `make_flags(plan_cost=True, cs=False)` |
| `make_hour(**kw)` | 创建 HourData | `make_hour(staff_id='EMP002')` |
| `create_sap_instance(...)` | 创建 Sap 实例（跳过 COM） | `create_sap_instance(mock_session=session)` |
| `setup_session_mock(session, text_returns)` | 配置 findById 返回值 | 见上方模板 |

所有工厂函数都接受 `**overrides`，只需传入要覆盖的字段，其余使用默认值。

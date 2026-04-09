"""
sap/test 共享 pytest fixtures

使用方法:
    pytest 自动加载 conftest.py，fixture 通过参数名注入测试函数。
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import (
    make_config, make_flags, create_sap_instance, SapResult,
)


@pytest.fixture
def mock_session():
    """Mock SAP GUI session"""
    return MagicMock()


@pytest.fixture
def sap_instance(mock_session):
    """创建 Sap 实例，跳过 __init__ COM 连接"""
    return create_sap_instance(mock_session=mock_session)

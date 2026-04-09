"""
测试 Sap.unlock_or_lock_order — 解锁/锁定订单

使用方法:
    pytest sap/test/test_unlock_or_lock_order.py -v

覆盖分支:
    - Unlock: 两个 checkbox 设为 False, KVGR4 = "100"
    - Lock: checkbox 设为 True, KVGR4 = " "
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance, setup_session_mock


class TestUnlockOrLockOrder:
    """unlock_or_lock_order 解锁/锁定"""

    def test_unlock_success(self):
        """Unlock — 返回成功，message 包含 'Unlock 成功'"""
        session = MagicMock()
        setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)

        result = sap.unlock_or_lock_order('Unlock')

        assert result.success
        assert result.message == 'Unlock 成功'

    def test_lock_success(self):
        """Lock — 返回成功，message 包含 'Lock 成功'"""
        session = MagicMock()
        setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)

        result = sap.unlock_or_lock_order('Lock')

        assert result.success
        assert result.message == 'Lock 成功'

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.unlock_or_lock_order('Unlock')

        assert not result.success
        assert 'Unlock 未成功' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

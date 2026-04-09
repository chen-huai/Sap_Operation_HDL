"""
测试 Sap.vf01_operate — 添加形式发票 (VF01)

使用方法:
    pytest sap/test/test_vf01_operate.py -v
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance


class TestVf01Operate:
    """vf01_operate 形式发票创建"""

    def test_success(self):
        """正常执行 — 返回成功 SapResult"""
        session = MagicMock()
        sap = create_sap_instance(mock_session=session)

        result = sap.vf01_operate()

        assert result.success
        # 验证事务码和执行按钮
        calls = [str(c) for c in session.findById.call_args_list]
        assert any('okcd' in c for c in calls)
        assert any('btn[11]' in c for c in calls)

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("VF01 界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.vf01_operate()

        assert not result.success
        assert '形式发票添加失败' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

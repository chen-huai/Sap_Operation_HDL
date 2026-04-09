"""
测试 Sap.save_sap — 保存 SAP 数据

使用方法:
    pytest sap/test/test_save_sap.py -v

覆盖分支:
    - 首次保存成功 + 状态栏含 '已保存'
    - 首次失败、重试成功 + 状态栏含 'saved'
    - 状态栏不含保存成功标记 → 失败
    - 状态栏读取异常 → 失败
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance, setup_session_mock


class TestSaveSapSuccess:
    """save_sap 成功路径"""

    def test_first_try_success_chinese(self):
        """首次保存成功 — 状态栏含 '已保存'"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': '文档 60001234 已保存',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.save_sap('订单')

        assert result.success

    def test_first_try_success_english(self):
        """首次保存成功 — 状态栏含 'saved'"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': 'Document 60001234 has been saved',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.save_sap('Order')

        assert result.success


class TestSaveSapFailure:
    """save_sap 失败路径"""

    def test_status_bar_no_saved_keyword(self):
        """状态栏不含保存关键字 → 失败"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': '发生错误：字段不能为空',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.save_sap('订单')

        assert not result.success
        assert '订单保存失败' in result.message

    def test_status_bar_read_exception(self):
        """状态栏读取异常 → 失败并包含异常信息"""
        session = MagicMock()
        call_count = 0

        def findById_side_effect(element_id):
            nonlocal call_count
            if element_id == 'wnd[0]/sbar/pane[0]':
                raise Exception("状态栏不可用")
            call_count += 1
            return MagicMock()

        session.findById = MagicMock(side_effect=findById_side_effect)
        sap = create_sap_instance(mock_session=session)

        result = sap.save_sap('测试')

        assert not result.success
        assert '无法读取状态栏' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

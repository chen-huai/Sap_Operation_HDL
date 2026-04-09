"""
测试 Sap.save_hours — 保存工时数据

使用方法:
    pytest sap/test/test_save_hours.py -v

覆盖分支:
    - 首次保存成功 (btn[11] + OPTION1)
    - 首次失败 → 重试后状态栏含 'Data was saved'
    - 首次失败 → 重试后通过 btn[11] 保存
    - 超过最大重试次数 → 失败
"""

import pytest
from unittest.mock import MagicMock, call

from sap.test.helpers import create_sap_instance, setup_session_mock


class TestSaveHoursSuccess:
    """save_hours 成功路径"""

    def test_first_try_success(self):
        """首次 btn[11] + OPTION1 成功 — 无异常"""
        session = MagicMock()
        setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)

        result = sap.save_hours()

        assert result.success


class TestSaveHoursRetry:
    """save_hours 重试路径"""

    def test_retry_data_was_saved(self):
        """首次失败 → 重试后状态栏 'Data was saved'"""
        session = MagicMock()
        call_count = 0

        def findById_side_effect(element_id):
            nonlocal call_count
            call_count += 1
            el = MagicMock()
            # 首次 btn[11] 成功, OPTION1 抛异常
            if call_count == 2:
                raise Exception("OPTION1 not found")
            # 重试阶段: 状态栏返回 'Data was saved'
            if element_id == 'wnd[0]/sbar/pane[0]':
                el.text = 'Data was saved'
            return el

        session.findById = MagicMock(side_effect=findById_side_effect)
        sap = create_sap_instance(mock_session=session)

        result = sap.save_hours()

        assert result.success
        assert result.message == '录Hour成功'

    def test_retry_with_button_save(self):
        """首次失败 → 重试后通过 btn[11] 保存"""
        session = MagicMock()
        call_count = 0

        def findById_side_effect(element_id):
            nonlocal call_count
            call_count += 1
            el = MagicMock()
            # 触发首次异常
            if call_count == 2:
                raise Exception("OPTION1 not found")
            # 重试: 状态栏返回其他文本，走 else 分支
            if element_id == 'wnd[0]/sbar/pane[0]':
                el.text = 'Ready'
            return el

        session.findById = MagicMock(side_effect=findById_side_effect)
        sap = create_sap_instance(mock_session=session)

        result = sap.save_hours()

        assert result.success


class TestSaveHoursMaxRetries:
    """save_hours 超限失败"""

    def test_exceed_max_retries(self):
        """所有重试均失败 → 返回失败结果"""
        session = MagicMock()
        call_count = 0

        def findById_side_effect(element_id):
            nonlocal call_count
            call_count += 1
            # 首次 OPTION1 失败
            if call_count == 2:
                raise Exception("首次失败")
            # 所有重试都失败
            raise Exception("重试失败")

        session.findById = MagicMock(side_effect=findById_side_effect)
        sap = create_sap_instance(mock_session=session)

        result = sap.save_hours()

        assert not result.success
        assert '已重试14次' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

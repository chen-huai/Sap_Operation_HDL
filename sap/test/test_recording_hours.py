"""
测试 Sap.recording_hours — 在工时表中记录工时

使用方法:
    pytest sap/test/test_recording_hours.py -v

覆盖分支:
    - 首行空 → 直接在 row 0 填写
    - 前几行非空 → 跳过，在首个空行填写
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance, make_hour


class TestRecordingHours:
    """recording_hours 工时记录"""

    def test_first_row_empty(self):
        """首行空 — 在 row 0 填写"""
        session = MagicMock()
        elements = {}

        def findById(element_id):
            if element_id not in elements:
                el = MagicMock()
                # 行检测字段: 首行返回空字符串
                if 'txtZRUCKDS-DATUMK[2,0]' in element_id:
                    el.text = ''
                elements[element_id] = el
            return elements[element_id]

        session.findById = MagicMock(side_effect=findById)
        sap = create_sap_instance(mock_session=session)
        hour = make_hour()

        result = sap.recording_hours(hour)

        assert result.success

    def test_skip_nonempty_rows(self):
        """前两行非空 — 跳过，在 row 2 填写"""
        session = MagicMock()
        elements = {}

        def findById(element_id):
            if element_id not in elements:
                el = MagicMock()
                if 'txtZRUCKDS-DATUMK[2,0]' in element_id:
                    el.text = '2026.04.01'  # 非空
                elif 'txtZRUCKDS-DATUMK[2,1]' in element_id:
                    el.text = '2026.04.02'  # 非空
                elif 'txtZRUCKDS-DATUMK[2,2]' in element_id:
                    el.text = ''  # 空 → 在此行填写
                elements[element_id] = el
            return elements[element_id]

        session.findById = MagicMock(side_effect=findById)
        sap = create_sap_instance(mock_session=session)
        hour = make_hour()

        result = sap.recording_hours(hour)

        assert result.success

    def test_with_explicit_row_num(self):
        """显式指定起始行号"""
        session = MagicMock()
        elements = {}

        def findById(element_id):
            if element_id not in elements:
                el = MagicMock()
                if 'txtZRUCKDS-DATUMK[2,5]' in element_id:
                    el.text = ''
                elements[element_id] = el
            return elements[element_id]

        session.findById = MagicMock(side_effect=findById)
        sap = create_sap_instance(mock_session=session)

        result = sap.recording_hours(make_hour(), row_num=5)

        assert result.success

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("工时表异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.recording_hours(make_hour())

        assert not result.success
        assert '录Hour失败' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

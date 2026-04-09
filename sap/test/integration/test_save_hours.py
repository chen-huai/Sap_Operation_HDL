"""
集成测试: save_hours 保存工时

前提条件:
    - SAP GUI 已登录
    - 已录入工时（recording_hours 之后）
    - test_config.py: ALLOW_SAVE = True

运行: pytest sap/test/integration/test_save_hours.py -v -s
"""

import pytest


class TestSaveHours:
    def test_save_hours(self, sap_live, require_save):
        """保存工时数据"""
        result = sap_live.save_hours()
        assert result.success, f'工时保存失败: {result.message}'
        print(f'工时保存成功: {result.message}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

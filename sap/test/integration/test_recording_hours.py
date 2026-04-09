"""
集成测试: recording_hours 记录工时

前提条件:
    - SAP GUI 已登录
    - 已通过 login_hour_gui 进入工时界面
    - test_config.py: HOUR 已填写

运行: pytest sap/test/integration/test_recording_hours.py -v -s
"""

import pytest


class TestRecordingHours:
    def test_recording(self, sap_live, hour):
        """在工时表中记录工时"""
        result = sap_live.recording_hours(hour)
        assert result.success, f'工时记录失败: {result.message}'
        print(f'工时记录成功, 订单: {hour.order_no}, 工时: {hour.allocated_hours}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

"""
集成测试: login_hour_gui 登录工时 GUI (ZRU1)

前提条件:
    - SAP GUI 已登录
    - test_config.py: HOUR.staff_id, HOUR.week 已填写

运行: pytest sap/test/integration/test_login_hour_gui.py -v -s
"""

import pytest


class TestLoginHourGui:
    def test_login_hour(self, sap_live, hour):
        """登录工时系统"""
        result = sap_live.login_hour_gui(hour)
        assert result.success, f'工时登录失败: {result.message}'
        print(f'工时系统登录成功, 员工: {hour.staff_id}, 周次: {hour.week}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

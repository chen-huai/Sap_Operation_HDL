"""
集成测试: VF01 添加形式发票

前提条件:
    - SAP GUI 已登录
    - 当前订单已有 Item（通常在 VA02 + save 之后）

运行: pytest sap/test/integration/test_vf01_operate.py -v -s
"""

import pytest


class TestVf01Operate:
    def test_vf01(self, sap_live):
        """VF01 添加形式发票"""
        result = sap_live.vf01_operate()
        assert result.success, f'VF01 失败: {result.message}'
        print(f'VF01 执行成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

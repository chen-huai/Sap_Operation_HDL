"""
集成测试: VF03 查看形式发票

前提条件:
    - SAP GUI 已登录
    - 已有形式发票（通常在 VF01 + save 之后）

运行: pytest sap/test/integration/test_vf03_operate.py -v -s
"""

import pytest


class TestVf03Operate:
    def test_vf03(self, sap_live):
        """VF03 查看形式发票 — 返回 proforma_no"""
        result = sap_live.vf03_operate()
        assert result.success, f'VF03 失败: {result.message}'
        print(f'VF03 执行成功, Proforma No: {result.proforma_no}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

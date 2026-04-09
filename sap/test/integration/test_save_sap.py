"""
集成测试: save_sap 保存数据

前提条件:
    - SAP GUI 已登录
    - 当前有待保存的操作（VA01/VA02 等之后）
    - test_config.py: ALLOW_SAVE = True

运行: pytest sap/test/integration/test_save_sap.py -v -s
"""

import pytest


class TestSaveSap:
    def test_save(self, sap_live, require_save):
        """保存当前 SAP 操作"""
        result = sap_live.save_sap('集成测试')
        assert result.success, f'保存失败: {result.message}'
        print(f'保存成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])

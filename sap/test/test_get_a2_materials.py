"""
测试 Sap._get_a2_materials — A2 物料映射查找

使用方法:
    pytest sap/test/test_get_a2_materials.py -v
"""

import pytest

from sap.test.helpers import create_sap_instance


class TestGetA2Materials:
    """_get_a2_materials 纯逻辑测试（无 COM 依赖）"""

    def test_match_405(self):
        """物料代码含 '405' → 返回 405 映射"""
        sap = create_sap_instance()
        result = sap._get_a2_materials('A2-405-00')
        assert result == ('T75-405-00', 'T20-405-00')

    def test_match_430(self):
        """物料代码含 '430' → 返回 430 映射"""
        sap = create_sap_instance()
        result = sap._get_a2_materials('A2-430-00')
        assert result == ('T20-430-00', 'T75-430-00')

    def test_match_441(self):
        """物料代码含 '441' → 返回 441 映射"""
        sap = create_sap_instance()
        result = sap._get_a2_materials('A2-441-00')
        assert result == ('T75-441-00', 'T20-441-00')

    def test_no_match_defaults_to_441(self):
        """物料代码无匹配 → 默认回退到 441 映射"""
        sap = create_sap_instance()
        result = sap._get_a2_materials('A2-999-00')
        assert result == ('T75-441-00', 'T20-441-00')


if __name__ == '__main__':
    pytest.main([__file__, '-v'])

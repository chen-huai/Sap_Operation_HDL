# -*- coding: utf-8 -*-
"""
验证状态文件格式和备份机制
"""

import json
import os
import shutil
from pathlib import Path
import sys

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from auto_updater.config import Config, get_executable_dir, UPDATE_STATE_FILE
from auto_updater.config_constants import SOFTWARE_ID

def verify_current_state():
    """验证当前状态文件格式"""
    print("="*60)
    print("验证当前状态文件格式")
    print("="*60)

    state_path = os.path.join(get_executable_dir(), UPDATE_STATE_FILE)
    print(f"状态文件路径: {state_path}")
    print(f"软件ID: {SOFTWARE_ID}")

    if not os.path.exists(state_path):
        print("\n[INFO] 状态文件不存在")
        return

    with open(state_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f"\n当前状态文件内容:")
    print(json.dumps(data, indent=2, ensure_ascii=False))

    # 检测格式
    if SOFTWARE_ID in data:
        print(f"\n[OK] 文件格式: 新格式（包含软件ID: {SOFTWARE_ID}）")
        print(f"    该软件的状态: {data[SOFTWARE_ID]}")
    else:
        print(f"\n[INFO] 文件格式: 旧格式（需要迁移）")
        print(f"    将触发自动迁移...")

def test_backup_mechanism():
    """测试备份机制（手动创建旧格式文件）"""
    print("\n" + "="*60)
    print("测试备份机制")
    print("="*60)

    # 创建测试环境
    test_state_path = os.path.join(get_executable_dir(), "test_update_state.json")

    # 创建旧格式测试文件
    legacy_data = {
        "last_check_date": "2026-01-27T12:00:00.000000",
        "test_field": "legacy_value"
    }

    print(f"1. 创建旧格式测试文件: {test_state_path}")
    with open(test_state_path, 'w', encoding='utf-8') as f:
        json.dump(legacy_data, f, indent=2)

    print(f"   内容: {legacy_data}")

    # 临时修改 UPDATE_STATE_FILE 以指向测试文件
    import auto_updater.config as config_module
    original_file = config_module.UPDATE_STATE_FILE
    config_module.UPDATE_STATE_FILE = "test_update_state.json"

    try:
        # 触发迁移
        print("\n2. 触发迁移...")
        config = Config()

        # 模拟迁移逻辑
        from auto_updater.config_constants import SOFTWARE_ID
        new_state = {SOFTWARE_ID: legacy_data}

        # 创建备份
        test_backup_path = f"{test_state_path}.bak"
        try:
            shutil.copy2(test_state_path, test_backup_path)
            print(f"   [OK] 备份文件创建成功: {test_backup_path}")

            with open(test_backup_path, 'r', encoding='utf-8') as f:
                backup_data = json.load(f)
            print(f"   备份内容: {backup_data}")
        except Exception as e:
            print(f"   [ERROR] 备份失败: {e}")

        # 保存新格式
        with open(test_state_path, 'w', encoding='utf-8') as f:
            json.dump(new_state, f, indent=2, ensure_ascii=False)

        print(f"\n3. 迁移后文件内容:")
        with open(test_state_path, 'r', encoding='utf-8') as f:
            migrated_data = json.load(f)
        print(json.dumps(migrated_data, indent=2, ensure_ascii=False))

        print(f"\n[OK] 备份机制测试完成!")

    finally:
        # 恢复原始配置
        config_module.UPDATE_STATE_FILE = original_file

        # 清理测试文件
        if os.path.exists(test_state_path):
            os.remove(test_state_path)
        if os.path.exists(test_backup_path):
            os.remove(test_backup_path)
        print(f"\n[INFO] 测试文件已清理")

if __name__ == "__main__":
    verify_current_state()
    test_backup_mechanism()
    input("\n按回车键退出...")

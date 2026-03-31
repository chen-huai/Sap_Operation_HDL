# -*- coding: utf-8 -*-
"""
通用打包+数字签名脚本
======================
复用方式：将 code_signer/ 文件夹和本文件一起复制到新项目，
修改下方「项目配置」区域即可。

使用方法：
    python build_with_signing.py
"""

import os
import sys
import subprocess
import shutil
import glob
import time
from pathlib import Path

# 导入签名模块（来自 code_signer 包）
try:
    sys.path.insert(0, str(Path(__file__).parent))
    from code_signer.sign_exe_file import sign_exe_with_sha1, verify_exe_signature
    SIGNER_AVAILABLE = True
except ImportError as err:
    SIGNER_AVAILABLE = False
    print(f"[警告] 签名模块不可用: {err}")
    def sign_exe_with_sha1(exe_path): return False, "签名模块不可用"
    def verify_exe_signature(exe_path): return False, "签名模块不可用"


# ====================== 项目配置（复用时修改此处）======================
CONFIG = {
    'main_script': 'Sap_Operate_theme.py',   # 主程序文件
    'icon_file':   'ch.ico',                  # 图标文件（.ico）
    'exe_name':    'Sap_Operate_theme',        # 输出 EXE 名称（不含扩展名）
    'console':     False,                      # True=显示控制台，False=窗口模式
}


# ====================== 工具函数 ======================

def print_header(text):
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)


def print_step(step_num, text):
    print(f"\n[步骤 {step_num}] {text}")
    print("-" * 60)


def check_files():
    """检查必要文件是否存在"""
    print_step(1, "检查必要文件")

    missing = []
    for f in [CONFIG['main_script'], CONFIG['icon_file']]:
        exists = os.path.exists(f)
        print(f"  {'✓' if exists else '✗'} {f}")
        if not exists:
            missing.append(f)

    print(f"  {'✓' if SIGNER_AVAILABLE else '⚠'} code_signer: "
          f"{'可用' if SIGNER_AVAILABLE else '不可用，将跳过签名'}")

    if missing:
        print(f"\n❌ 缺少必要文件: {', '.join(missing)}")
        return False

    print("\n✅ 文件检查通过")
    return True


def clean_build_artifacts():
    """清理旧的打包文件"""
    print_step(2, "清理旧的打包文件")

    removed = []
    for d in ['build', 'dist']:
        if os.path.exists(d):
            try:
                shutil.rmtree(d)
                print(f"  ✓ 删除: {d}/")
                removed.append(d)
            except Exception as e:
                print(f"  ✗ 删除失败 {d}: {e}")

    for spec in glob.glob(f"{CONFIG['exe_name']}.spec"):
        try:
            os.remove(spec)
            print(f"  ✓ 删除: {spec}")
            removed.append(spec)
        except Exception as e:
            print(f"  ✗ 删除失败 {spec}: {e}")

    print(f"\n✅ 已清理 {len(removed)} 项" if removed else "\n✅ 无需清理")


def build_exe():
    """使用 PyInstaller 打包 EXE（标准 spec 配置）"""
    print_step(3, "开始打包")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed" if not CONFIG['console'] else "--console",
        "--clean",
        "--noconfirm",
        f"--icon={CONFIG['icon_file']}",
        f"--name={CONFIG['exe_name']}",
        CONFIG['main_script'],
    ]

    print(f"  执行: PyInstaller {CONFIG['main_script']}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            exe_path = f"dist/{CONFIG['exe_name']}.exe"
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"\n✅ 打包成功!")
                print(f"  文件: {exe_path}  ({size_mb:.1f} MB)")
                return True, exe_path
            return False, "打包完成但找不到 EXE 文件"
        else:
            print("\n❌ 打包失败!")
            for line in result.stderr.split('\n')[-20:]:
                if line.strip():
                    print(f"    {line}")
            return False, result.stderr

    except Exception as e:
        print(f"\n❌ 打包异常: {e}")
        return False, str(e)


def main():
    print_header("自动打包 + 数字签名")
    print(f"  主程序: {CONFIG['main_script']}")
    print(f"  图标:   {CONFIG['icon_file']}")
    print(f"  签名:   {'code_signer' if SIGNER_AVAILABLE else '不可用'}")

    start_time = time.time()

    try:
        if not check_files():
            input("\n按回车退出...")
            return False

        clean_build_artifacts()

        success, result = build_exe()
        if not success:
            print(f"\n❌ 打包失败: {result}")
            input("\n按回车退出...")
            return False

        exe_path = result

        # 签名
        print_step(4, "数字签名")
        sign_success, sign_message = False, "签名模块不可用"
        if SIGNER_AVAILABLE:
            sign_success, sign_message = sign_exe_with_sha1(exe_path)
            if sign_success:
                verify_exe_signature(exe_path)
        else:
            print("  ⚠ 跳过签名步骤")

        # 完成报告
        elapsed = time.time() - start_time
        print_header("打包完成!")
        print(f"  文件: {exe_path}")
        print(f"  大小: {os.path.getsize(exe_path) / (1024*1024):.1f} MB")
        print(f"  签名: {'✓ 已签名' if sign_success else f'✗ 未签名 ({sign_message})'}")
        print(f"  耗时: {elapsed:.1f} 秒")
        print("=" * 60)

        try:
            if input("\n是否打开 dist 目录? (y/n): ").strip().lower() in ['y', 'yes', '是']:
                os.startfile(os.path.dirname(os.path.abspath(exe_path)))
        except Exception:
            pass

        input("\n按回车退出...")
        return True

    except KeyboardInterrupt:
        print("\n\n⚠ 用户取消操作")
        input("\n按回车退出...")
        return False
    except Exception as e:
        import traceback
        print(f"\n\n❌ 发生错误: {e}")
        traceback.print_exc()
        input("\n按回车退出...")
        return False


if __name__ == "__main__":
    main()

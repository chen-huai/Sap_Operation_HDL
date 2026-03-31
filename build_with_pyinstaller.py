#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAP工具打包脚本（简化版）
使用 PyInstaller 将 Python 应用程序打包为 Windows 可执行文件

更新日志（2025-12-25）：
- 激进简化：删除所有不必要的诊断和验证功能
- 只保留核心打包功能
- 代码从 804 行简化到约 250 行
- 不再需要 pandas 诊断、DLL 验证等冗余功能
"""

import os
import sys
import subprocess
import shutil
import time
from pathlib import Path
from typing import List


class PackagerConfig:
    """打包配置类"""

    def __init__(self):
        self.entry_file = 'Sap_Operate_theme.py'
        self.spec_file = 'Sap_Operate_theme.spec'
        self.icon_file = 'Sap_Operate_Logo.ico'
        self.app_name = 'Sap_Operate_theme'
        self.clean_build = True
        self.onefile = True  # 单文件模式


class SAPPackager:
    """SAP工具打包器（简化版）"""

    def __init__(self, config: PackagerConfig = None):
        self.config = config or PackagerConfig()
        self.start_time = time.time()

    def log(self, message: str, level: str = "INFO"):
        """日志输出"""
        timestamp = time.strftime("%H:%M:%S")
        # 替换特殊字符以避免编码问题
        safe_message = message.replace('✓', '[OK]').replace('✗', '[FAIL]').replace('-', '[SKIP]')
        print(f"[{timestamp}] {level}: {safe_message}")

    def check_environment(self) -> bool:
        """检查打包环境"""
        self.log("检查打包环境...")

        # 检查 Python 版本
        if sys.version_info < (3, 8):
            self.log("Python 版本过低，建议使用 3.8 或更高版本", "ERROR")
            return False

        # 检查 PyInstaller
        try:
            result = subprocess.run(['pyinstaller', '--version'],
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                self.log(f"PyInstaller 版本: {result.stdout.strip()}")
                return True
            else:
                self.log("PyInstaller 未安装", "ERROR")
                return False
        except (subprocess.TimeoutExpired, FileNotFoundError):
            self.log("PyInstaller 未安装", "ERROR")
            return False

    def clean_directories(self):
        """清理构建目录"""
        self.log("清理构建目录...")

        dirs_to_clean = ['build', 'dist']

        for dir_name in dirs_to_clean:
            if os.path.exists(dir_name):
                try:
                    shutil.rmtree(dir_name)
                    self.log(f"[OK] 清理 {dir_name}")
                except Exception as e:
                    self.log(f"[FAIL] 清理 {dir_name} 失败: {e}", "WARNING")

    def get_build_command(self) -> List[str]:
        """构建 PyInstaller 命令"""
        cmd = ['pyinstaller']

        # 优先使用 .spec 文件
        if os.path.exists(self.config.spec_file):
            cmd.extend([self.config.spec_file, '--clean', '--noconfirm'])
            self.log(f"使用配置文件: {self.config.spec_file}")
        else:
            # 使用命令行参数
            if self.config.onefile:
                cmd.append('--onefile')
            else:
                cmd.append('--onedir')

            cmd.append('--windowed')
            cmd.append('--clean')
            cmd.append('--noconfirm')

            # 添加图标
            if os.path.exists(self.config.icon_file):
                cmd.append(f'--icon={self.config.icon_file}')

            cmd.append(self.config.entry_file)
            self.log("使用命令行参数配置")

        return cmd

    def run_build(self) -> bool:
        """执行打包构建"""
        self.log("开始打包构建...")

        cmd = self.get_build_command()
        self.log(f"执行命令: {' '.join(cmd)}")

        try:
            # 设置环境变量
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'

            # 执行打包命令
            result = subprocess.run(
                cmd,
                env=env,
                check=True,
                text=True
            )

            self.log("[OK] 打包构建成功")
            return True

        except subprocess.CalledProcessError as e:
            self.log(f"[FAIL] 打包构建失败，返回码: {e.returncode}", "ERROR")
            self.log("提示: 查看上方的错误信息，通常是因为缺少依赖", "WARNING")
            return False

        except Exception as e:
            self.log(f"[FAIL] 打包过程中发生错误: {e}", "ERROR")
            return False

    def verify_build(self) -> bool:
        """验证打包结果"""
        self.log("验证打包结果...")

        exe_path = Path('dist') / f"{self.config.app_name}.exe"

        if exe_path.exists():
            file_size = exe_path.stat().st_size
            self.log(f"[OK] 可执行文件已生成: {exe_path}")
            self.log(f"[OK] 文件大小: {file_size / (1024*1024):.1f} MB")
            return True
        else:
            self.log("[FAIL] 可执行文件未生成", "ERROR")
            return False

    def show_summary(self):
        """显示打包摘要"""
        elapsed_time = time.time() - self.start_time
        self.log("=" * 50)
        self.log("打包摘要")
        self.log("=" * 50)
        self.log(f"[OK] 打包完成，耗时: {elapsed_time:.1f} 秒")
        self.log(f"[OK] 可执行文件位置: dist/{self.config.app_name}.exe")

        exe_path = Path('dist') / f"{self.config.app_name}.exe"
        if exe_path.exists():
            file_size = exe_path.stat().st_size
            self.log(f"[OK] 文件大小: {file_size / (1024*1024):.1f} MB")

        self.log("")
        self.log("使用说明:")
        self.log("1. 双击运行可执行文件")
        self.log("2. 如有问题，请检查系统环境")
        self.log("3. 确保已安装运行时库（如 Microsoft Visual C++）")
        self.log("=" * 50)

    def run(self) -> bool:
        """执行完整打包流程"""
        self.log("开始 SAP 工具打包流程")
        self.log("=" * 50)

        # 检查环境
        if not self.check_environment():
            return False

        # 清理目录
        if self.config.clean_build:
            self.clean_directories()

        # 执行构建
        if not self.run_build():
            return False

        # 验证结果
        if not self.verify_build():
            return False

        # 显示摘要
        self.show_summary()
        return True


def main():
    """主函数"""
    # 创建配置
    config = PackagerConfig()

    # 解析命令行参数
    if len(sys.argv) > 1:
        if '--no-clean' in sys.argv:
            config.clean_build = False
        if '--onedir' in sys.argv:
            config.onefile = False
        if '--help' in sys.argv or '-h' in sys.argv:
            print("SAP工具打包脚本（简化版）")
            print("用法: python build_with_pyinstaller.py [选项]")
            print("选项:")
            print("  --no-clean  不清理构建目录")
            print("  --onedir    生成目录模式（非单文件）")
            print("  --help      显示此帮助信息")
            return

    # 创建打包器并运行
    packager = SAPPackager(config)
    success = packager.run()

    if success:
        print("\n[SUCCESS] 打包成功！")
    else:
        print("\n[FAIL] 打包失败！")
        sys.exit(1)


if __name__ == '__main__':
    main()

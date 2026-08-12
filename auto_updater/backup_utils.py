# -*- coding: utf-8 -*-
"""
备份工具
在原文件旁创建带时间戳的副本，供更新失败时人工回退使用。

设计说明：
    生产环境走两阶段更新（新版本先落在 downloads/，主目录 exe 在替换成功前
    始终原样在位），本身已具备天然回退能力，因此这里只需最轻量的就地备份，
    不需要独立的备份目录管理与自动回滚机制。
"""

import os
import shutil
import time
import logging
from typing import Optional

logger = logging.getLogger(__name__)


def create_local_backup(file_path: str) -> Optional[str]:
    """
    在原文件旁创建 `<原文件名>.backup.<时间戳>` 副本

    Args:
        file_path: 要备份的文件路径

    Returns:
        备份文件路径；源文件不存在或备份失败时返回 None
    """
    try:
        if not file_path or not os.path.exists(file_path):
            logger.warning(f"[备份] 源文件不存在，跳过备份: {file_path}")
            return None

        backup_path = f"{file_path}.backup.{int(time.time())}"
        shutil.copy2(file_path, backup_path)
        logger.info(f"[备份] 已创建备份: {os.path.basename(backup_path)}")
        return backup_path

    except Exception as e:
        logger.warning(f"[备份] 创建备份失败: {e}")
        return None

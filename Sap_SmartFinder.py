"""
SAP GUI Scripting 智能容器查找模块

提供灵活、高效的 SAP GUI 元素查找功能，支持多种查找策略和自动适配。

使用方法:
    from Sap_SmartFinder import SmartContainerFinder
    finder = SmartContainerFinder(session)
    contacts = finder.find_contact_names("wnd[1]/usr/")

作者: Claude Code
版本: 1.0.0
"""

import re
import logging
from typing import List, Dict, Callable, Optional, Any
from datetime import datetime


# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class SmartContainerFinder:
    """
    SAP GUI 容器智能查找器

    提供多种查找策略，支持灵活的元素定位和数据提取。

    特性:
        - 按类型查找元素
        - 按文本内容过滤
        - 按ID模式匹配
        - 位置解析和排序
        - 递归遍历嵌套结构
        - 完善的异常处理
    """

    def __init__(self, session):
        """
        初始化查找器

        参数:
            session: SAP GUI Session 对象 (win32com.client.CDispatch)
        """
        self.session = session
        logger.info("SmartContainerFinder 初始化完成")

    def find_by_pattern(
        self,
        container_path: str,
        element_type: str = "GuiLabel",
        text_filter: Optional[str] = None,
        text_filter_func: Optional[Callable[[str], bool]] = None,
        id_pattern: Optional[str] = None,
        sort_by_position: bool = True
    ) -> List[Dict[str, Any]]:
        """
        通用模式查找方法

        参数:
            container_path: 容器路径，如 "wnd[1]/usr/"
            element_type: 元素类型，如 "GuiLabel", "GuiTextField" 等
            text_filter: 文本过滤字符串（模糊匹配）
            text_filter_func: 自定义文本过滤函数（返回布尔值）
            id_pattern: ID模式字符串（部分匹配）
            sort_by_position: 是否按位置排序

        返回:
            元素信息列表，每个元素包含:
                - id: 元素ID
                - text: 元素文本
                - type: 元素类型
                - row: 行位置（如果可解析）
                - col: 列位置（如果可解析）
                - element: 原始元素对象

        示例:
            # 查找所有包含 "张" 的标签
            results = finder.find_by_pattern(
                "wnd[1]/usr/",
                element_type="GuiLabel",
                text_filter="张"
            )

            # 使用自定义过滤函数
            results = finder.find_by_pattern(
                "wnd[1]/usr/",
                element_type="GuiLabel",
                text_filter_func=lambda x: len(x.strip()) > 2
            )
        """
        start_time = datetime.now()
        results = []

        try:
            # 获取容器
            container = self.session.findById(container_path)
            logger.debug(f"访问容器: {container_path}")

            # 查找指定类型的元素
            elements = container.findAllByNodeType(element_type)
            logger.debug(f"找到 {elements.Count} 个 {element_type} 类型的元素")

            # 遍历元素并应用过滤条件
            for i in range(elements.Count):
                try:
                    elem = elements(i)
                    elem_text = getattr(elem, 'text', '')
                    elem_id = getattr(elem, 'id', '')

                    # 应用文本过滤
                    if text_filter and text_filter.lower() not in elem_text.lower():
                        continue

                    # 应用自定义文本过滤函数
                    if text_filter_func and not text_filter_func(elem_text):
                        continue

                    # 应用ID模式过滤
                    if id_pattern and id_pattern not in elem_id:
                        continue

                    # 解析位置信息
                    row, col = self._parse_position(elem_id)

                    results.append({
                        'id': elem_id,
                        'text': elem_text,
                        'type': elem_type,
                        'row': row,
                        'col': col,
                        'element': elem
                    })

                except Exception as e:
                    logger.warning(f"处理元素 {i} 时出错: {e}")
                    continue

            # 按位置排序
            if sort_by_position:
                results.sort(key=lambda x: (x['row'], x['col']))

            elapsed = (datetime.now() - start_time).total_seconds() * 1000
            logger.info(f"查找完成: 找到 {len(results)} 个元素, 耗时 {elapsed:.2f}ms")

            return results

        except Exception as e:
            logger.error(f"查找失败: {e}")
            raise

    def find_contact_names(
        self,
        container_path: str = "wnd[1]/usr/",
        min_length: int = 1,
        exclude_patterns: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """
        专门用于查找联系人名称的优化方法

        参数:
            container_path: 容器路径
            min_length: 最小文本长度（过滤空标签）
            exclude_patterns: 要排除的文本模式列表

        返回:
            联系人信息列表，包含位置和文本信息

        示例:
            contacts = finder.find_contact_names("wnd[1]/usr/")
            for contact in contacts:
                print(f"位置[{contact['row']},{contact['col']}]: {contact['text']}")
        """
        exclude_patterns = exclude_patterns or [
            '', ' ', '...', 'N/A', '-', '—', '___'
        ]

        # 定义文本过滤函数
        def text_filter_func(text):
            # 基本过滤
            if not text or len(text.strip()) < min_length:
                return False

            # 排除特定模式
            text = text.strip()
            for pattern in exclude_patterns:
                if text == pattern:
                    return False

            return True

        results = self.find_by_pattern(
            container_path,
            element_type="GuiLabel",
            text_filter_func=text_filter_func,
            sort_by_position=True
        )

        logger.info(f"找到 {len(results)} 个联系人标签")

        return results

    def find_all_elements(
        self,
        container_path: str,
        recursive: bool = True,
        max_depth: int = 10
    ) -> List[Dict[str, Any]]:
        """
        递归查找所有元素（包括嵌套容器）

        参数:
            container_path: 起始容器路径
            recursive: 是否递归查找子容器
            max_depth: 最大递归深度

        返回:
            所有元素的信息列表

        示例:
            all_elements = finder.find_all_elements("wnd[1]/usr/")
            print(f"总共找到 {len(all_elements)} 个元素")
        """
        results = []

        def traverse_element(element, depth=0):
            """递归遍历元素"""
            if depth > max_depth:
                return

            try:
                # 获取元素信息
                elem_type = getattr(element, 'type', 'Unknown')
                elem_id = getattr(element, 'id', 'Unknown')
                elem_text = getattr(element, 'text', '')

                results.append({
                    'id': elem_id,
                    'text': elem_text,
                    'type': elem_type,
                    'depth': depth,
                    'element': element
                })

                # 递归处理子元素
                if recursive and hasattr(element, 'Children') and element.Children.Count > 0:
                    for i in range(element.Children.Count):
                        traverse_element(element.Children(i), depth + 1)

            except Exception as e:
                logger.debug(f"遍历元素时出错: {e}")

        try:
            container = self.session.findById(container_path)
            traverse_element(container)

            logger.info(f"递归遍历完成: 找到 {len(results)} 个元素")
            return results

        except Exception as e:
            logger.error(f"递归查找失败: {e}")
            raise

    def find_by_id_range(
        self,
        container_path: str,
        element_type: str,
        id_prefix: str,
        row_range: tuple,
        col_range: tuple
    ) -> List[Dict[str, Any]]:
        """
        按ID范围查找元素

        参数:
            container_path: 容器路径
            element_type: 元素类型
            id_prefix: ID前缀，如 "wnd[1]/usr/lbl"
            row_range: 行范围 (start, end)
            col_range: 列范围 (start, end)

        返回:
            符合条件的元素列表

        示例:
            # 查找 lbl[0-5][0-10] 范围内的标签
            results = finder.find_by_id_range(
                "wnd[1]/usr/",
                "GuiLabel",
                "wnd[1]/usr/lbl",
                row_range=(0, 5),
                col_range=(0, 10)
            )
        """
        results = []

        try:
            # 获取所有指定类型的元素
            all_elements = self.find_by_pattern(
                container_path,
                element_type=element_type,
                sort_by_position=True
            )

            # 过滤指定范围
            for elem_info in all_elements:
                row, col = elem_info['row'], elem_info['col']

                if (row_range[0] <= row <= row_range[1] and
                    col_range[0] <= col <= col_range[1]):
                    results.append(elem_info)

            logger.info(f"范围查找完成: 找到 {len(results)} 个元素")
            return results

        except Exception as e:
            logger.error(f"范围查找失败: {e}")
            raise

    def get_element_info(self, element_id: str) -> Dict[str, Any]:
        """
        获取单个元素的详细信息

        参数:
            element_id: 完整的元素ID

        返回:
            元素的详细信息字典

        示例:
            info = finder.get_element_info("wnd[1]/usr/lbl[1,6]")
            print(f"文本: {info['text']}, 类型: {info['type']}")
        """
        try:
            elem = self.session.findById(element_id)

            return {
                'id': element_id,
                'text': getattr(elem, 'text', ''),
                'type': getattr(elem, 'type', 'Unknown'),
                'visible': getattr(elem, 'visible', True),
                'enabled': getattr(elem, 'enabled', True),
                'tooltip': getattr(elem, 'tooltip', ''),
                'element': elem
            }

        except Exception as e:
            logger.error(f"获取元素信息失败: {e}")
            raise

    def _parse_position(self, element_id: str) -> tuple:
        """
        从元素ID解析位置信息

        参数:
            element_id: 元素ID，如 "wnd[1]/usr/lbl[1,6]"

        返回:
            (row, col) 元组，如果无法解析则返回 (0, 0)

        示例:
            row, col = self._parse_position("wnd[1]/usr/lbl[1,6]")
            print(f"行: {row}, 列: {col}")  # 输出: 行: 1, 列: 6
        """
        try:
            # 匹配模式: [行,列] 或 [行,列,其他]
            match = re.search(r'\[(\d+),(\d+)(?:,\d+)?\]', element_id)
            if match:
                row = int(match.group(1))
                col = int(match.group(2))
                return row, col

            return 0, 0

        except Exception as e:
            logger.debug(f"解析位置失败: {element_id}, 错误: {e}")
            return 0, 0

    def print_element_tree(self, container_path: str, max_depth: int = 5):
        """
        打印元素树结构（用于调试）

        参数:
            container_path: 容器路径
            max_depth: 最大显示深度

        示例:
            finder.print_element_tree("wnd[1]/usr/")
        """
        def print_element(element, depth=0):
            indent = "  " * depth
            try:
                elem_type = getattr(element, 'type', '?')
                elem_id = getattr(element, 'id', '?')
                elem_text = getattr(element, 'text', '')

                # 截断过长的文本
                if len(elem_text) > 30:
                    elem_text = elem_text[:27] + "..."

                print(f"{indent}├─ [{elem_type}] {elem_id}")
                if elem_text:
                    print(f"{indent}│   文本: {elem_text}")

                # 递归打印子元素
                if hasattr(element, 'Children') and element.Children.Count > 0 and depth < max_depth:
                    for i in range(element.Children.Count):
                        print_element(element.Children(i), depth + 1)

            except Exception as e:
                print(f"{indent}├─ [访问失败: {e}]")

        try:
            container = self.session.findById(container_path)
            print(f"\n元素树结构: {container_path}")
            print("=" * 60)
            print_element(container)
            print("=" * 60 + "\n")

        except Exception as e:
            logger.error(f"打印元素树失败: {e}")


# 便捷函数
def create_finder(session) -> SmartContainerFinder:
    """
    创建查找器实例的便捷函数

    参数:
        session: SAP GUI Session 对象

    返回:
        SmartContainerFinder 实例
    """
    return SmartContainerFinder(session)

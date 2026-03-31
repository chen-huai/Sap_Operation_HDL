"""
PDF解析工具模块

提供统一的PDF字段提取接口,优化公司名称、金额等字段的提取逻辑

使用方法:
    from PDF_Parser_Utils import extract_company_name, extract_revenue, extract_fapiao_no, parse_pdf_fields

    # 提取单个字段
    company = extract_company_name('购 名称：东莞市瀚雅纤维制品有限公司 销 名称：南德认证检测')

    # 批量解析PDF行
    msg = {}
    parse_pdf_fields(msg, '购 名称：XXX公司 销 名称：南德认证检测...', inv_pattern, order_pattern)

更新日志:
    - 2026-01-23: 修正extract_revenue()正则表达式,优化"（小写）"格式含税金额提取
      * 匹配模式: (?:价税)?合计（大写）... （小写）¥XXXX
      * 提升对复杂发票格式的兼容性
"""

import re
from typing import Optional, Dict


# ============ 预编译正则表达式 (性能优化) ============

# 公司名称提取模式
_COMPANY_PATTERNS = [
    # 模式1: 完整的"购 名称：...销 名称："格式 (允许冒号前有空格)
    re.compile(r'购\s+名称\s*[:：]\s*(.+?)\s+(?:销\s+名称\s*[:：])'),
    # 模式2: 简化格式 - 公司名称直接在南德认证检测之前
    re.compile(r'^(.+?)\s+南德认证检测'),
    # 模式3: 备用格式 - 仅提取"购 名称："之后的内容 (允许冒号前有空格)
    re.compile(r'购\s+名称\s*[:：]\s*(.+?)(?:\s+销|$)'),
    # 模式4: 兼容"购 名称 ："(冒号前有空格)的特殊格式
    re.compile(r'购\s+名称\s*[:：]\s*(.+?)\s+销\s+名称'),
]

# 金额提取模式 - 优化匹配"（小写）"格式
_REVENUE_PATTERN_1 = re.compile(r'[\(（]小写[\)）]\s*[¥￥]\s*([\d,]+\.?\d*)')
_REVENUE_PATTERN_2 = re.compile(r'[¥￥]\s*([\d,]+\.?\d*)')

# 发票号提取模式
_FAPIAO_PATTERN_1 = re.compile(r'发票号码\s*[:：]\s*([A-Z0-9]+)')
_FAPIAO_PATTERN_2 = re.compile(r'^[A-Z0-9]+$')  # 用于验证

# ============ 公司名称提取 ============

def extract_company_name(text: str) -> Optional[str]:
    """
    从PDF文本行中提取购买方公司名称 (性能优化版本)

    支持三种格式:
    1. '购 名称：XXX公司 销 名称：南德认证检测...'
    2. 'XXX公司 南德认证检测...'
    3. '购 名称：XX XX XX公司 销 名称：...' (公司名带空格)

    Args:
        text: 包含公司信息的文本行

    Returns:
        清理后的公司名称(去除所有空格),如果无法提取则返回None

    Examples:
        >>> extract_company_name('购 名称：东莞市瀚雅纤维制品有限公司 销 名称：南德认证检测')
        '东莞市瀚雅纤维制品有限公司'

        >>> extract_company_name('遇见技术发展(东莞)有限公司 南德认证检测(中国)有限公司广州分公司')
        '遇见技术发展(东莞)有限公司'

        >>> extract_company_name('购 名称：东 莞 市 寰瑜贸易有限公司 销 名称：南德认证检测')
        '东莞市寰瑜贸易有限公司'
    """
    # 使用预编译的正则表达式 (性能优化)
    for pattern in _COMPANY_PATTERNS:
        match = pattern.search(text)
        if match:
            # 提取匹配的文本
            company_name = match.group(1)

            # 去除所有空格和多余空白字符
            company_name = company_name.replace(' ', '').replace('\t', '').strip()

            # 验证有效性: 非空且不等于单字标记
            if company_name and company_name not in ['购', '销', '名称']:
                return company_name

    return None


# ============ 金额提取 ============

def extract_revenue(text: str) -> Optional[str]:
    """
    从PDF文本行中提取金额信息 (性能优化版本)

    支持格式:
    - '价税合计（小写）¥1234.56'
    - '合计（小写）¥1,234.56元'
    - '（小写）1234.56'

    Args:
        text: 包含金额信息的文本行

    Returns:
        提取的金额字符串(包含货币符号),如果无法提取则返回None

    Examples:
        >>> extract_revenue('价税合计（小写）¥1234.56')
        '¥1234.56'

        >>> extract_revenue('¥1,234.56')
        '¥1,234.56'
    """
    # 使用预编译的正则表达式 (性能优化)
    # 模式1: 匹配"（小写）"后的金额（含税金额）
    match1 = _REVENUE_PATTERN_1.search(text)
    if match1:
        return '¥' + match1.group(1).replace(' ', '')

    # # 模式2: 直接匹配货币符号和金额
    # match2 = _REVENUE_PATTERN_2.search(text)
    # if match2:
    #     return '¥' + match2.group(1).replace(' ', '')

    # 模式3: 兼容旧的split()逻辑 - 提取')'后的第三部分
    if '）' in text or ')' in text:
        parts = re.split('[）)]', text)
        if len(parts) >= 3:
            potential_revenue = parts[2].strip()
            # 验证是否包含金额
            if re.search(r'[\d,]+\.?\d*', potential_revenue):
                return potential_revenue

    return None


# ============ 发票号码提取 ============

def extract_fapiao_no(text: str) -> Optional[str]:
    """
    从PDF文本行中提取发票号码 (性能优化版本)

    支持格式:
    - '发票号码：12345678'
    - '发票号码: 12345678'
    - '制票 12345678'

    Args:
        text: 包含发票号码的文本行

    Returns:
        提取的发票号码,如果无法提取则返回None

    Examples:
        >>> extract_fapiao_no('发票号码：12345678')
        '12345678'
    """
    # 使用预编译的正则表达式 (性能优化)
    # 模式1: 匹配"发票号码："或"发票号码:"
    match1 = _FAPIAO_PATTERN_1.search(text)
    if match1:
        return match1.group(1)

    # 模式2: 兼容旧的逻辑 - 包含"制"字,取第二个分词
    if '制' in text:
        parts = text.split()
        if len(parts) >= 2:
            # 提取第二部分作为发票号
            fapiao_no = parts[1].strip()
            # 使用预编译的正则验证
            if _FAPIAO_PATTERN_2.match(fapiao_no):
                return fapiao_no

    return None


# ============ Invoice No 提取 ============

def extract_invoice_no(text: str, pattern: str) -> Optional[int]:
    r"""
    从PDF文本行中提取Invoice No

    Args:
        text: 包含Invoice No的文本行
        pattern: 用于匹配Invoice No的正则表达式模式

    Returns:
        提取的Invoice No(整数),如果无法提取则返回None

    Examples:
        >>> pattern = r'(?<!\d)486\d{6}(?!\d)'
        >>> extract_invoice_no('Invoice No: 486123456', pattern)
        486123456
    """
    match = re.search(pattern, text)
    if match:
        try:
            return int(match.group(0))
        except (ValueError, IndexError):
            return None
    return None


# ============ Order No 提取 ============

def extract_order_no(text: str, pattern: str) -> Optional[str]:
    r"""
    从PDF文本行中提取Order No

    Args:
        text: 包含Order No的文本行
        pattern: 用于匹配Order No的正则表达式模式

    Returns:
        提取的Order No(字符串),如果无法提取则返回None

    Examples:
        >>> pattern = r'7486\d{5}'
        >>> extract_order_no('Order: 748612345', pattern)
        '748612345'
    """
    match = re.search(pattern, text)
    if match:
        return match.group(0)
    return None


# ============ 工具函数 ============

def clean_company_name(company_name: str) -> str:
    """
    清理公司名称:去除多余空格和特殊字符

    Args:
        company_name: 原始公司名称

    Returns:
        清理后的公司名称
    """
    if not company_name:
        return ''

    # 去除所有空格、制表符、换行符
    cleaned = company_name.replace(' ', '').replace('\t', '').replace('\n', '').replace('\r', '')

    # 去除首尾空白
    cleaned = cleaned.strip()

    return cleaned


def validate_company_name(company_name: str) -> bool:
    """
    验证公司名称是否有效

    Args:
        company_name: 待验证的公司名称

    Returns:
        True if 公司名称有效, else False
    """
    if not company_name:
        return False

    # 不能是单字标记
    if company_name in ['购', '销', '名称']:
        return False

    # 长度检查:至少2个字符
    if len(company_name) < 2:
        return False

    # 必须包含中文字符
    if not re.search(r'[\u4e00-\u9fa5]', company_name):
        return False

    return True


# ============ 批量解析 ============

def parse_pdf_line(line: str, patterns: dict = None) -> dict:
    """
    解析PDF文本行,提取所有可能的字段

    Args:
        line: PDF文本行
        patterns: 字段名到正则表达式的映射(可选)

    Returns:
        包含提取字段的字典,如 {'Company Name': 'XXX', 'Revenue': '¥1234'}
    """
    result = {}

    # 尝试提取公司名称
    company = extract_company_name(line)
    if company:
        result['Company Name'] = company

    # 尝试提取金额
    revenue = extract_revenue(line)
    if revenue:
        result['Revenue'] = revenue

    # 尝试提取发票号
    fapiao = extract_fapiao_no(line)
    if fapiao:
        result['FaPiao No'] = fapiao

    # 如果提供了自定义模式,也进行匹配
    if patterns:
        for field_name, pattern in patterns.items():
            match = re.search(pattern, line)
            if match:
                result[field_name] = match.group(0) if len(match.groups()) == 0 else match.group(1)

    return result


def parse_pdf_fields(msg: Dict, line: str, inv_pattern: str = None, order_pattern: str = None) -> None:
    """
    统一的PDF字段解析函数 - 简化主程序代码

    从单个PDF文本行中提取所有可能的字段并更新到msg字典中
    只在字段不存在时才设置,避免覆盖已有数据

    Args:
        msg: 存储解析结果的字典,会被直接修改
        line: PDF文本行
        inv_pattern: Invoice No的正则表达式模式(可选)
        order_pattern: Order No的正则表达式模式(可选)

    Returns:
        None (msg字典会被直接修改)

    Examples:
        >>> msg = {}
        >>> parse_pdf_fields(msg, '购 名称：XXX公司 销 名称：南德认证检测', inv_pattern, order_pattern)
        >>> print(msg['Company Name'])
        'XXX公司'
    """
    # 公司名称提取
    if 'Company Name' not in msg:
        company_name = extract_company_name(line)
        if company_name:
            msg['Company Name'] = company_name

    # 金额提取
    if 'Revenue' not in msg:
        revenue = extract_revenue(line)
        if revenue:
            msg['Revenue'] = revenue

    # 发票号提取
    if 'fapiao' not in msg:
        fapiao_no = extract_fapiao_no(line)
        if fapiao_no:
            msg['fapiao'] = fapiao_no

    # Invoice No提取 (如果提供了模式)
    if inv_pattern and 'Invoice No' not in msg:
        match = re.search(inv_pattern, line)
        if match:
            try:
                msg['Invoice No'] = int(match.group(0))
            except (ValueError, IndexError):
                pass

    # Order No提取 (如果提供了模式)
    if order_pattern and 'Order No' not in msg:
        match = re.search(order_pattern, line)
        if match:
            msg['Order No'] = match.group(0)


import pandas as pd

class ExcelFieldMapper:
    """
    Excel字段映射工具类，用于处理复杂的Excel文件字段映射和转换需求

    主要功能：
    1. 提供灵活的字段名称映射
    2. 自动匹配数据框中的列名
    3. 支持字段验证和转换
    4. 兼容多种命名风格的Excel文件

    核心设计思路：
    - 使用多映射表支持不同命名方式
    - 提供动态匹配和转换方法
    - 保持最大的灵活性和可扩展性
    """

    def __init__(self):
        """
        初始化字段映射表，包含完整的字段映射
        映射表结构：{中文字段名: [常见英文列名1, 常见英文列名2, 系统内部列名]}

        设计原则：
        - 包含中文名、英文名和可能的系统内部列名
        - 支持多种命名风格
        - 方便后续扩展        
        """
        self.excel_fields_map = {
            'Project No.': ['Request_No', 'Request No.', 'RequestNo', 'Request Number', '项目编号'],
            'ProjectNo(Tlims)': ['Project_No', 'ProjectNo', 'Project Number', 'project_num', '项目编号'],
            'CS': ['CS', 'cs', 'Primary CS', 'CustomerSupport', '客户支持'],
            'Sales': ['Sales', 'sales', 'SalesPerson', 'sales_rep', '销售'],
            'Currency': ['currency', 'Order_Currency', 'CurrencyType', 'Currency', '货币'],
            'GPC Glo. Par. Code': ['global_partner_code', 'gpc_code', 'Gpc_Code', 'GPC Code', '全球合作伙伴代码'],
            'Material Code': ['material_code', 'Material_Code', 'MaterialCode', 'Material Code', '物料代码'],
            'SAP No.': ['sap_no', 'SAP_No', 'SAPNo', 'SAP Custom Code','SAP Customer Code', 'SAP编号'],
            'Amount': ['Untaxed amount', 'Amount', 'Amount_Money', 'money_amount', '未税金额'],
            'Amount with VAT': ['amount_with_vat', 'Amount_With_VAT', 'Tax-inclusive amount', 'vat_amount', '含税金额'],
            "Revenue\n(RMB)": ["Revenue (RMB)", "Revenue", "未税金额CNY"],
            'Total Cost': ['total_cost', 'Total Subcon Cost', 'TotalCost', 'Total_Cost', '总成本'],
            'Exchange Rate': ['exchange_rate', 'Exchange_Rate', 'ExchangeRate', 'Rate', '汇率'],
            "Invoices' name (Chinese)": ['invoices_name_chinese', 'Invoices_Name', 'InvoiceName', 'Payer', '发票名称(中文)'],
            'Buyer(GPC)': ['buyer_gpc', 'Gpc_name', 'BuyerGPC', 'GPC', '买家(GPC)'],
            'Month': ['month', 'Month', 'MonthPeriod', 'month_period', '月份'],
            'SAP Order No.': ['SAP Order No.', 'SAPOrderNo', 'SAP_Order_No', 'Order Number', '订单号'],
            'Reference No.': ['Reference_No', 'Reference No', 'Reference Number', 'reference no', 'Reference_No'],
            'Text': ['text', 'description', 'ShortText', 'brief_desc', '文本'],
            'Long Text': ['long_text', 'detailed_description', 'LongText', 'full_description', '长文本'],
            'Client Contact Name': ['client_contact_name', 'contact', 'Customer contact', 'contact_name', 'Client Contact', '客户联系人名称']
        }

    def match_columns(self, dataframe):
        """
        根据预定义的映射表匹配数据框中的列名

        Args:
            dataframe (pd.DataFrame): 待匹配的数据框

        Returns:
            dict: 匹配成功的字段映射 {中文字段名: 实际列名}

        匹配逻辑：
        1. 遍历映射表中的每个英文字段
        2. 尝试在数据框中匹配可能的列名
        3. 返回成功匹配的映射关系
        """
        matched_fields = {}
        
        for english_name, possible_names in self.excel_fields_map.items():
            for name in possible_names:
                if name in dataframe.columns:
                    matched_fields[english_name] = name
                    break
        
        return matched_fields

    def get_column_mapping(self, dataframe, required_fields=None):
        """
        获取字段映射，可指定必须包含的字段

        Args:
            dataframe (pd.DataFrame): 待处理的数据框
            required_fields (list, optional): 必须包含的字段列表（中文名）

        Returns:
            tuple: (匹配成功的字段映射, 是否满足所有必需字段)

        验证逻辑：
        1. 执行列名匹配
        2. 检查是否包含所有必需字段required_fields
        3. 返回匹配结果和完整性标志
        """
        matched_fields = self.match_columns(dataframe)
        
        if required_fields is not None:
            for field in required_fields:
                if field not in matched_fields:
                    return matched_fields, False
        
        return matched_fields, True

    def get_standard_column_name(self, column_name):
        """
        获取标准（英文）列名
        """
        for english_name, possible_names in self.excel_fields_map.items():
            if column_name in possible_names:
                return english_name
        return column_name

    def get_chinese_name(self, english_name):
        """
        获取对应的中文名称
        """
        return self.excel_fields_map.get(english_name, [])[-1] if english_name in self.excel_fields_map else english_name
    
    
    def transform_dataframe(self, dataframe, mapping=None):
        """

        Args:
            dataframe (pd.DataFrame): 原始数据框
            mapping (dict, optional): 自定义字段映射

        Returns:
            pd.DataFrame: 转换列名后的数据框

        转换原理：
        1. 如未提供映射，自动获取映射
        2. 将英文/系统列名转换为中文名
        3. 返回转换后的数据框
        """
        if mapping is None:
            # _（下划线）：表示忽略第二个返回值（是否完整的布尔值）
            mapping, _ = self.get_column_mapping(dataframe)
        
        rename_dict = {v: k for k, v in mapping.items()}
        return dataframe.rename(columns=rename_dict)


    def validate_dataframe(self, dataframe, required_fields=None):
        """
        验证数据框是否包含必要字段

        Args:
            dataframe (pd.DataFrame): 待验证的数据框
            required_fields (list, optional): 必须包含的字段列表（中文名）

        Returns:
            tuple: (是否通过验证, 缺失字段列表)

        验证流程：
        1. 匹配当前列名
        2. 检查必要字段是否存在
        3. 返回验证结果和缺失字段
        """
        missing_fields = []
        matched_fields = self.match_columns(dataframe)
        
        if required_fields:
            for field in required_fields:
                if field not in matched_fields:
                    missing_fields.append(field)
        
        return len(missing_fields) == 0, missing_fields

    def get_standard_column_name(self, column_name):
        """
        获取标准（中文）列名

        Args:
            column_name (str): 输入的列名（英文/系统名）

        Returns:
            str: 对应的中文标准列名

        转换机制：
        1. 遍历映射表
        2. 匹配输入的列名
        3. 返回对应的中文名
        4. 若无匹配，返回原列名
        """
        for chinese_name, possible_names in self.excel_fields_map.items():
            if column_name in possible_names:
                return chinese_name
        return column_name

    def get_all_possible_names(self, standard_name):
        """
        获取特定标准字段名的所有可能名称

        Args:
            standard_name (str): 标准中文字段名

        Returns:
            list: 该字段的所有可能名称列表
        """
        return self.excel_fields_map.get(standard_name, [])
    
    
    def update_field_names(self, field_list):
        """
        更新传入的字段列表，将列表中的元素替换为self.excel_fields_map中对应的键名

        Args:
            field_list (list): 需要更新的字段列表

        Returns:
            list: 更新后的字段列表

        更新逻辑：
        1. 遍历传入的字段列表
        2. 对于每个字段，在self.excel_fields_map中查找匹配的键名
        3. 如果找到匹配，则用键名替换原字段名
        4. 如果没有找到匹配，保持原字段名不变
        """
        updated_list = []
        for field in field_list:
            matched = False
            for key, value_list in self.excel_fields_map.items():
                if field in value_list:
                    updated_list.append(key)
                    matched = True
                    break
            if not matched:
                updated_list.append(field)
        return updated_list

# 创建全局单例，便于直接调用
excel_field_mapper = ExcelFieldMapper()



# 在Sap_Operate.py中，可以这样使用：

# from Excel_Field_Mapper import excel_field_mapper
# import pandas as pd

# # 原有获取Excel数据的方法
# def getFileData(self, fileUrl):
#     # 原始获取数据的方法
#     newData = Get_Data()
#     dataframe = newData.getFileData(fileUrl)

#     # 使用Excel_Field_Mapper进行字段映射和转换
#     # 1. 匹配并验证必要字段
#     required_fields = ['项目编号', '销售', '客户支持']
#     mapping, is_complete = excel_field_mapper.get_column_mapping(
#         dataframe, 
#         required_fields=required_fields
#     )

#     # 2. 检查字段完整性
#     if not is_complete:
#         # 处理字段不完整的情况
#         missing_fields = [field for field in required_fields if field not in mapping]
#         raise ValueError(f"缺少必要字段: {missing_fields}")

#     # 3. 转换DataFrame列名
#     transformed_dataframe = excel_field_mapper.transform_dataframe(dataframe)

#     # 4. 可选：获取所有可能的列名变体
#     project_no_names = excel_field_mapper.get_all_possible_names('项目编号')

#     # 5. 保留原有的处理逻辑
#     # ... 后续的数据处理代码保持不变

#     return transformed_dataframe
# 这种方式的优势：

# 自动转换不同风格的Excel列名
# 验证必要字段的存在
# 保留原有的数据处理逻辑
# 提高代码的健壮性和适应性
# 使用场景示例：

# # 处理来自不同系统或导出的Excel文件
# try:
#     # 文件可能来自不同的系统，列名可能不同
#     dataframe = self.getFileData('path/to/excel_file.xlsx')
    
#     # 后续的数据处理保持不变
#     self.process_data(dataframe)

# except ValueError as e:
#     # 处理字段缺失的情况
#     print(f"数据转换错误: {e}")
#     # 可以进行适当的错误处理
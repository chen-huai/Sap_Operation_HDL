import pandas as pd
import numpy as np
from Excel_Field_Mapper import excel_field_mapper
import ast
class Get_Data():
    # def __init__(self,fileDataUrl):
    #     self.fileDataUrl = fileDataUrl
        # self.getFileData()
        # self.getHeaderData()
        # self.getIndexNumForHead()
        # self.getFileDataList()

    @staticmethod
    def _convert_datetime_to_str(fileData):
        """
        转换 DataFrame 中的 datetime 列为字符串，避免保存时类型冲突
        :param fileData: pandas DataFrame
        :return: 转换后的 DataFrame
        """
        datetime_cols = fileData.select_dtypes(include=['datetime64']).columns
        if len(datetime_cols) > 0:
            fileData[datetime_cols] = fileData[datetime_cols].astype(str)
        return fileData

    def getFileData(self, fileDataUrl):
        """
        读取文件数据并处理类型转换
        修复：强制 Client Contact Name 相关列在读取时为字符串类型，避免 float64 类型错误
        """
        self.fileDataUrl = fileDataUrl
        fileType = self.fileDataUrl.split(".")[-1]

        # 定义需要强制为字符串类型的列（Client Contact Name 的所有可能列名变体）
        contact_columns = [
            'Client Contact Name', 'client_contact_name', 'contact',
            'Customer contact', 'contact_name', 'Client Contact', '客户联系人名称'
        ]
        dtype_dict = {col: 'str' for col in contact_columns}

        if fileType == 'xlsx':
            self.fileData = pd.read_excel(self.fileDataUrl, dtype=dtype_dict)
            # self.fileData = pd.read_excel(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_excel(self.fileDataUrl, keep_default_na=False)
        elif fileType == 'csv':
            self.fileData = pd.read_csv(self.fileDataUrl, dtype=dtype_dict)
            # self.fileData = pd.read_csv(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_csv(self.fileDataUrl, keep_default_na=False)
        height, width = self.fileData.shape

        self.fileData = excel_field_mapper.transform_dataframe(self.fileData)
        self.fileData = self._convert_datetime_to_str(self.fileData)
        return self.fileData

    def getFileTableData(self, fileDataUrl):
        """
        读取文件数据用于表格显示
        修复：强制 Client Contact Name 相关列在读取时为字符串类型
        """
        self.fileDataUrl = fileDataUrl
        fileType = self.fileDataUrl.split(".")[-1]

        # 定义需要强制为字符串类型的列
        contact_columns = [
            'Client Contact Name', 'client_contact_name', 'contact',
            'Customer contact', 'contact_name', 'Client Contact', '客户联系人名称'
        ]
        dtype_dict = {col: 'str' for col in contact_columns}

        if fileType == 'xlsx':
            self.fileData = pd.read_excel(self.fileDataUrl, dtype=dtype_dict)
        elif fileType == 'csv':
            self.fileData = pd.read_csv(self.fileDataUrl, dtype=dtype_dict)
        height, width = self.fileData.shape

        self.fileData = self._convert_datetime_to_str(self.fileData)
        return self.fileData

    def getFileMoreSheetData(self, fileDataUrl, sheet_name=[]):
        """
        读取多sheet Excel文件并合并
        修复：强制 Client Contact Name 相关列在读取时为字符串类型
        """
        if sheet_name==[]:
            sheet_name = None
        self.fileDataUrl = fileDataUrl

        # 定义需要强制为字符串类型的列
        contact_columns = [
            'Client Contact Name', 'client_contact_name', 'contact',
            'Customer contact', 'contact_name', 'Client Contact', '客户联系人名称'
        ]
        dtype_dict = {col: 'str' for col in contact_columns}

        self.fileData = pd.read_excel(self.fileDataUrl, sheet_name=sheet_name, dtype=dtype_dict)
        self.fileData = pd.concat(self.fileData.values(), ignore_index=True)
        self.fileData.dropna(subset=['Final Invoice No.'], inplace=True)
        self.fileData = self._convert_datetime_to_str(self.fileData)
        return self.fileData

    def getMergeFileData(self, fileDataUrl):
        self.fileDataUrl = fileDataUrl
        fileType = self.fileDataUrl.split(".")[-1]
        if fileType == 'xlsx':
            # self.fileData = pd.read_excel(self.fileDataUrl)
            self.fileData = pd.read_excel(self.fileDataUrl, float_precision='round_trip', dtype='str')
            # self.fileData = pd.read_excel(self.fileDataUrl, keep_default_na=False)
        elif fileType == 'csv':
            # self.fileData = pd.read_csv(self.fileDataUrl)
            self.fileData = pd.read_csv(self.fileDataUrl, dtype='str')
            # self.fileData = pd.read_csv(self.fileDataUrl, keep_default_na=False)
        height, width = self.fileData.shape
        self.fileData = self._convert_datetime_to_str(self.fileData)
        return self.fileData
    def getHeaderData(self):
        self.headData = list(self.fileData.head())
        return self.headData
    def getIndexNumForHead(self):
        self.projectNo = self.headData.index('Project No.')
        self.cs = self.headData.index('CS')
        self.sales = self.headData.index('Sales')
        self.currency = self.headData.index('Currency')
        self.partnerCode = self.headData.index('GPC Glo. Par. Code')
        self.materialCode = self.headData.index('Material Code')
        self.phyMaterialCode = self.headData.index('PHY Material Code')
        self.chmMaterialCode = self.headData.index('CHM Material Code')
        self.sapNo = self.headData.index('SAP No.')
        self.amount = self.headData.index('Amount')
        self.amountWithVAT = self.headData.index('Amount with VAT')
        self.exchangeRate = self.headData.index('Exchange Rate')
        self.costList = list(self.fileData['Total Cost'])
        return self.projectNo, self.cs, self.sales, self.currency, self.partnerCode, self.materialCode, self.phyMaterialCode, self.chmMaterialCode, self.sapNo, self.amount, self.amountWithVAT, self.exchangeRate,self.costList
    def deleteTheRows(self, deleteRowList = {}):
        for key in deleteRowList:
            self.fileData = self.fileData[self.fileData[key] != deleteRowList[key]]
        return self.fileData
    def fillNanColumn(self,fillNanColumnKey):
        for filledKey in fillNanColumnKey:
            for fillKey in fillNanColumnKey[filledKey]:
                self.fileData[filledKey] = self.fileData[filledKey].fillna(self.fileData[fillKey])
        # self.fileData["Material Code"].fillna(self.fileData["PHY Material Code"], inplace=True)
        # self.fileData["Material Code"].fillna(self.fileData["CHM Material Code"], inplace=True)
        return self.fileData
    # def pivotTable(self):
    def pivotTable(self,pivotTableKey, valusKey):
        pivotData = pd.pivot_table(self.fileData, index=pivotTableKey, values=valusKey, aggfunc='sum')
        return pivotData
    def getFileDataList(self,getFileDataListKey):
        self.fileDataList = {}
        for each in getFileDataListKey:
            self.fileDataList[each] = list(self.fileData[each])
        return self.fileDataList

    def getFileDataList1(self):
        self.fileData = self.fileData[self.fileData['Amount'] != 0]
        # self.fileData.dropna(axis=0,subset=["Amount"],inplace = True)
        self.projectNoList = list(self.fileData['Project No.'])
        self.csList = list(self.fileData['CS'])
        self.salesList = list(self.fileData['Sales'])
        self.currencyList = list(self.fileData['Currency'])
        self.partnerCodeList = list(self.fileData['GPC Glo. Par. Code'])
        self.materialCodeList = list(self.fileData['Material Code'])
        self.sapNoList = list(self.fileData['SAP No.'])
        self.amountList = list(self.fileData['Amount'])
        self.amountWithVATList = list(self.fileData['Amount with VAT'])
        self.exchangeRateList = list(self.fileData['Exchange Rate'])
        self.costList = list(self.fileData['Total Cost'])
        return self.projectNoList, self.csList, self.salesList, self.currencyList, self.partnerCodeList, self.materialCodeList,self.sapNoList, self.amountList, self.amountWithVATList, self.exchangeRateList,self.costList

    def deleteTheColumn(self, deleteColumnList):
        self.fileData.drop(columns=deleteColumnList, inplace=True)
        return self.fileData

    def mergeData(self, data1, data2, onData):
        mergeData = pd.merge(data1, data2, on=onData, how='inner')
        return mergeData

    def column_concat_func(self, data):
        # 行信息合并
        return pd.Series({
                'combine_column_msg': '\n'.join(data['column_msg'].unique()),
            })

    def row_concat_func(self, data):
        # 列信息合并
        return pd.Series({
                'combine_row_msg': '\n'.join(data['row_msg'].unique()),
            }
            )

    def rename_hour_fields(self, data, required_fields=None):
        """
        重命名工时数据字段，如果字段不存在则创建空字段
        """
        # 定义所需的标准字段
        # required_fields = {
        #     'staff_id': 'staff_id',
        #     'week': 'week',
        #     'order_no': 'order_no',
        #     'allocated_hours': 'allocated_hours',
        #     'office_time': 'office_time',
        #     'material_code': 'material_code',
        #     'item': 'item',
        #     'allocated_day': 'allocated_day',
        #     'staff_name': 'staff_name',
        # }
        str_data = required_fields
        required_fields = ast.literal_eval(str_data)

        # 创建新的DataFrame，包含所有必需字段
        new_data = pd.DataFrame()

        # 遍历所需字段，如果原数据中有则重命名，没有则创建空列
        for old_field, new_field in required_fields.items():
            if old_field in data.columns:
                new_data[new_field] = data[old_field]
            else:
                new_data[new_field] = None

        # 保留原数据中的其他列
        for col in data.columns:
            if col not in required_fields:
                new_data[col] = data[col]

        # 使用pandas的sort_values方法直接排序
        if 'staff_id' in new_data.columns and 'week' in new_data.columns:
            new_data = new_data.sort_values(['staff_id', 'week'])

        return new_data

# data = {
#     'col1': [1, 2, 3],
#     'col2': [4, 5, 6],
#     'col3': [7, 8, 9],
#     'col4': [10, 11, 12]
# }
#
# df = pd.DataFrame(data)
#
# # 选择要合并的三列，并使用apply函数将它们相加并用制表符隔开
# # df['merged'] = df[['col1', 'col2', 'col3']].apply(lambda row: '\t'.join(map(str, row)), axis=1)
#
# df['merged'] = df[['col1', 'col3', 'col2']].apply(lambda row: '\n'.join(f"{col}:{val}" for col, val in zip(df[['col1', 'col3', 'col2']], row)), axis=1)
#
# # 移除原始三列
# df.drop(['col1', 'col2', 'col3'], axis=1, inplace=True)



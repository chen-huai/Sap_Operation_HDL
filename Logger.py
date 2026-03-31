import pandas as pd
import datetime
import os


class Logger:
    def __init__(self, log_file, columns):
        self.log_file = log_file
        self.columns = columns
        self.log_df = pd.DataFrame(columns=columns)

    def log(self, data):
        """
        记录日志数据
        :param data: 字典类型，键必须与columns中的列名（除了'Update'）匹配
        """
        # 检查数据字典的键是否与列名匹配（除了'Update'列）
        expected_keys = set(col for col in self.columns if col != 'Update')
        actual_keys = set(data.keys())
        
        if not actual_keys.issubset(expected_keys):
            missing_keys = expected_keys - actual_keys
            extra_keys = actual_keys - expected_keys
            error_msg = []
            if missing_keys:
                error_msg.append(f"Missing keys: {missing_keys}")
            if extra_keys:
                error_msg.append(f"Extra keys: {extra_keys}")
            raise ValueError(f"Data keys do not match columns: {', '.join(error_msg)}")

        # 添加时间戳
        timestamp = datetime.datetime.now()
        log_data = {'Update': timestamp, **data}
        
        # 确保所有必需的列都有值
        for col in self.columns:
            if col not in log_data:
                log_data[col] = None

        # 添加新行
        self.log_df.loc[len(self.log_df)] = log_data

    def save_log_to_excel(self):
        """
        保存日志到Excel文件
        """
        # 转换 datetime 列为字符串，避免类型冲突
        datetime_cols = self.log_df.select_dtypes(include=['datetime64']).columns
        if len(datetime_cols) > 0:
            self.log_df[datetime_cols] = self.log_df[datetime_cols].astype(str)

        try:
            self.log_df.to_excel(self.log_file, index=False, merge_cells=False)
            print(f"Log saved successfully to {self.log_file}")
        except Exception as e:
            print(f"Error saving log: {e}")
            # 尝试保存到临时目录
            try:
                temp_dir = os.path.join(os.path.expanduser("~"), "temp")
                os.makedirs(temp_dir, exist_ok=True)
                temp_file = os.path.join(temp_dir, os.path.basename(self.log_file))
                self.log_df.to_excel(temp_file, index=False, merge_cells=False)
                print(f"Log saved to temporary location: {temp_file}")
            except Exception as e:
                print(f"Failed to save log to temporary location: {e}")

# # 创建Logger对象，传递列名作为参数
# logger = Logger("log.csv", ["Timestamp", "Message", "Value"])
#
# # 记录日志
# logger.log(["This is a log message", 42])
# logger.log(["Another log message", 123])
#
# # 保存日志到CSV文件
# logger.save_log_to_csv()

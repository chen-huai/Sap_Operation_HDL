import sys
import os
import re
import time
import math
import pandas as pd
import csv
import copy
import numpy as np
import win32com.client
import datetime
import shutil
import logging

from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction, QLabel
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon, QFontDatabase
from Get_Data import *
from PDF_Parser_Utils import extract_company_name, extract_revenue, extract_fapiao_no, parse_pdf_fields, PDF_Operate
from Sap_Function import *
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *
from auto_updater.config_constants import CURRENT_VERSION
from auto_updater import AutoUpdater, UI_AVAILABLE
from sap import CostOptions, OrderData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession
from runtime_globals import configContent, myWin, staff_dict

class HourMixin:
    def get_hour_file_url(self, position):
        fileUrl = myWin.getFile(configContent['Hour_Files_Import_URL'])
        if fileUrl:
            position.setText(fileUrl)
            QApplication.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
    def get_hour_combine_file(self):
        fileUrl = self.lineEdit_30.text()
        pivot_table_key = self.lineEdit_39.text().split(';')
        if fileUrl and pivot_table_key:
            try:
                self.textBrowser_4.append("数据开始合并")
                QApplication.processEvents()
                newData = Get_Data()
                file_data = newData.getFileTableData(fileUrl)
                # 删除
                deleteRowList = {'Order Number': ''}
                newData.deleteTheRows(deleteRowList)
                # 合并
                valus_key = ['Revenue', 'Total Subcon Cost']
                pivot_table_data = newData.pivotTable(pivot_table_key, valus_key)
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
                pivot_table_data_path = '%s\\%s' % (configContent['Hour_Files_Export_URL'], '1.order data %s.xlsx' % current_time)
                pivot_table_data_file = pivot_table_data.to_excel(pivot_table_data_path, merge_cells=False)
                self.lineEdit_37.setText(pivot_table_data_path)
                self.textBrowser_4.append("合并完成")
                self.textBrowser_4.append("文件路径：%s" % pivot_table_data_path)
            except Exception as errorMsg:
                self.textBrowser_4.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                QApplication.processEvents()
        elif pivot_table_key == []:
            self.textBrowser_4.append("请输入合并的key")
        else:
            self.textBrowser_4.append("请重新选择ODM文件")
    def update_config_content(self, update_data):
        # 创建配置字典的深拷贝以避免污染原始配置
        config_content = copy.deepcopy(configContent)
        config_content.update(update_data)
        return config_content  # 返回修改后的副本
    def get_department_hour(self):
        """
        计算部门工时并保存结果
        """
        order_data_path = self.lineEdit_37.text()
        hour_gui_data = myWin.getHourGuiData()
        if order_data_path:
            self.textBrowser_4.append("部门开始计算")
            QApplication.processEvents()
            
            # 更新配置内容
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取订单数据
            order_data_obj = Get_Data()
            order_datas = order_data_obj.getFileTableData(order_data_path)
            
            # 初始化结果DataFrame
            all_results = []
            
            # 调用hour方法处理每个订单
            revenue_allocator_obj = RevenueAllocator()
            for _, order_data in order_datas.iterrows():
                # 将Series转换为字典
                order_dict = order_data.to_dict()
                # 计算部门工时
                order_revenue_data = revenue_allocator_obj.allocate_department_hours(order_dict, config_content)
                all_results.extend(order_revenue_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            # material_code包含D2或D3，更新字段item=1000
            mask = result_df['material_code'].str.contains(r'D[23]', case=False, na=False, regex=True)
            result_df.loc[mask, 'item'] = '1000'
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            dept_hour_path = f"{configContent['Hour_Files_Export_URL']}\\2.dept hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(dept_hour_path, index=False)
            
            # 更新UI
            self.lineEdit_38.setText(dept_hour_path)
            self.textBrowser_4.append("部门计算完成")
            self.textBrowser_4.append(f"文件路径：{dept_hour_path}")
            QApplication.processEvents()
        else:
            self.textBrowser_4.append("请重新选择合并后的文件")
            QApplication.processEvents()
    def get_person_hour(self):
        """
        分配人员工时并保存结果
        """
        dept_hour_path = self.lineEdit_38.text()
        if dept_hour_path:
            self.textBrowser_4.append("开始分配人员")
            QApplication.processEvents()
            
            # 获取配置数据
            hour_gui_data = myWin.getHourGuiData()
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取参数
            max_hours_per_day = int(config_content['Max_Hour'])
            start_date = datetime.datetime.strptime(config_content['Start_Date'], '%Y.%m.%d').date()
            end_date = datetime.datetime.strptime(config_content['End_Date'], '%Y.%m.%d').date()
            
            # 获取部门工时数据
            dept_hour_obj = Get_Data()
            dept_hour_datas = dept_hour_obj.getFileTableData(dept_hour_path)
            
            # 计算各部门总工时
            dept_total_hours = dept_hour_datas.groupby('dept')['dept_hours'].sum().to_dict()
            
            # 初始化结果列表
            all_results = []
            
            # 处理每个部门工时记录
            revenue_allocator_obj = RevenueAllocator()
            for _, dept_hour in dept_hour_datas.iterrows():
                # 将Series转换为字典
                dept_hour_dict = dept_hour.to_dict()
                # 将单个记录转换为列表形式
                dept_hour_list = [dept_hour_dict]
                self.textBrowser_4.append(f"处理Order Number：{dept_hour_dict['order_no']}")
                QApplication.processEvents()
                # 分配人员工时
                person_hour_data = revenue_allocator_obj.allocate_person_average_hours(
                    dept_hour_list,
                    max_hours_per_day, 
                    start_date, 
                    end_date, 
                    staff_dict,
                    dept_total_hours,  # 添加部门总工时参数
                    config_content
                )
                all_results.extend(person_hour_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            person_hour_path = f"{configContent['Hour_Files_Export_URL']}\\3.person hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(person_hour_path, index=True, index_label='ID')
            
            # 打开结果文件
            os.startfile(person_hour_path)
            
            # 更新UI
            self.lineEdit_31.setText(person_hour_path)
            self.textBrowser_4.append("分配人员完成")
            self.textBrowser_4.append(f"文件路径：{person_hour_path}")
            QApplication.processEvents()
        else:
            self.textBrowser_4.append("请先完成部门工时计算")
            QApplication.processEvents()
    def get_average_person_hour(self):
        """
        使用平均分配方式分配人员工时并保存结果
        """
        dept_hour_path = self.lineEdit_38.text()
        if dept_hour_path:
            self.textBrowser_4.append("开始平均分配人员")
            QApplication.processEvents()

            # 获取配置数据
            hour_gui_data = myWin.getHourGuiData()
            config_content = self.update_config_content(hour_gui_data)

            # 获取参数
            max_hours_per_day = int(config_content['Max_Hour'])
            start_date = datetime.datetime.strptime(config_content['Start_Date'], '%Y.%m.%d').date()
            end_date = datetime.datetime.strptime(config_content['End_Date'], '%Y.%m.%d').date()

            # 获取部门工时数据
            dept_hour_obj = Get_Data()
            dept_hour_datas = dept_hour_obj.getFileTableData(dept_hour_path)

            # 计算各部门总工时
            dept_total_hours = dept_hour_datas.groupby('dept')['dept_hours'].sum().to_dict()
            
            # 初始化结果列表
            all_results = []
            
            # 按部门处理工时记录
            revenue_allocator_obj = RevenueAllocator()
            for dept, total_hours in dept_total_hours.items():
                self.textBrowser_4.append(f"\n处理部门：{dept}")
                self.textBrowser_4.append(f"部门总工时：{total_hours}")
                QApplication.processEvents()
                
                # 获取该部门的所有记录
                dept_records = dept_hour_datas[dept_hour_datas['dept'] == dept].to_dict('records')
                
                # 分配该部门的工时
                person_hour_data = revenue_allocator_obj.allocate_person_average_hours(
                    dept_records,
                    max_hours_per_day,
                    start_date,
                    end_date,
                    staff_dict,
                    {dept: total_hours},  # 只传入当前部门的总工时
                    config_content
                )
                all_results.extend(person_hour_data)
                
                # 显示分配结果
                allocated_hours = sum(record['allocated_hours'] for record in person_hour_data)
                allocation_rate = (allocated_hours / total_hours * 100) if total_hours > 0 else 0
                self.textBrowser_4.append(f"已分配工时：{allocated_hours:.2f}")
                self.textBrowser_4.append(f"分配率：{allocation_rate:.2f}%")
                QApplication.processEvents()

            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            person_hour_path = f"{configContent['Hour_Files_Export_URL']}\\3.person hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(person_hour_path, index=True, index_label='ID')
            
            # 打开结果文件
            os.startfile(person_hour_path)
            
            # 更新UI
            self.lineEdit_31.setText(person_hour_path)
            self.textBrowser_4.append("\n分配人员完成")
            self.textBrowser_4.append(f"文件路径：{person_hour_path}")
            
            # 显示总体分配统计
            total_original = sum(dept_total_hours.values())
            total_allocated = result_df['allocated_hours'].sum()
            allocation_rate = (total_allocated / total_original * 100) if total_original > 0 else 0
            
            self.textBrowser_4.append(f"\n总体分配统计:")
            self.textBrowser_4.append(f"原始总工时: {total_original:.2f}")
            self.textBrowser_4.append(f"已分配工时: {total_allocated:.2f}")
            self.textBrowser_4.append(f"分配率: {allocation_rate:.2f}%")
            
            QApplication.processEvents()
        else:
            self.textBrowser_4.append("请先完成部门工时计算")
            QApplication.processEvents()
    def clear_hour_gui(self):
        self.lineEdit_30.clear()
        self.lineEdit_37.clear()
        self.lineEdit_38.clear()
        self.lineEdit_31.clear()
        self.textBrowser_4.clear()
    def hourOperate(self):
        """
        处理工时数据并进行SAP操作
        """

        time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(configContent['Hour_Files_Export_URL'], f'log_{time_str}.xlsx')
        columns = [
            'ID',
            'staff_id',
            'week',
            'order_no',
            'allocated_hours',
            'office_time',
            'material_code',
            'item',
            'allocated_day',
            'staff_name',
            'status',
            'message',
            'Update'
        ]
        log_obj = Logger(log_file=log_file, columns=columns)
        try:
            # 获取文件路径
            hour_path = self.lineEdit_31.text()
            if not hour_path:
                QMessageBox.warning(self, "警告", "请先选择工时文件！")
                return



            # 获取并处理数据
            get_data = Get_Data()
            raw_data = get_data.getFileTableData(hour_path)

            # 重命名字段
            renamed_data = get_data.rename_hour_fields(raw_data, configContent['Hour_Field_Mapping'])

            # 初始化SAP操作对象
            sap = Sap()

            # 记录当前处理的staff_id和week
            current_staff_id = None
            current_week = None
            is_first_login = True  # 标记是否是第一次登录
            num = 0
            # 遍历分组后的数据
            for _, row in renamed_data.iterrows():
                num += 1
                staff_id = row['staff_id']
                week = row['week']
                log_data = {
                    'ID': '',
                    'staff_id': '',
                    'week': '',
                    'order_no': '',
                    'allocated_hours': '',
                    'office_time': '',
                    'material_code': '',
                    'item': '',
                    'allocated_day': '',
                    'staff_name': '',
                    'status': '',
                    'message': '',
                }  # 用于存储日志数据

                # 如果staff_id或week发生变化，需要重新登录
                if staff_id != current_staff_id or week != current_week:
                    # 如果不是第一次登录，需要先保存之前的工时
                    if not is_first_login:
                        save_res = sap.save_hours()
                        if not save_res['flag']:
                            error_msg = f"保存工时失败！Staff ID: {current_staff_id}, Week: {current_week}"
                            # logger.error(error_msg)
                            log_data.update({
                                'ID': row['ID'],
                                'staff_id': current_staff_id,
                                'week': current_week,
                                'order_no': row['order_no'],
                                'allocated_hours': row['allocated_hours'],
                                'office_time': row['office_time'],
                                'material_code': row['material_code'],
                                'item': row['item'],
                                'allocated_day': row['allocated_day'],
                                'staff_name': row['staff_name'],
                                'status': 'Failed',
                                'message': error_msg
                            })

                    # 登录SAP
                    login_res = sap.login_hour_gui(row)
                    if not login_res['flag']:
                        # logger.error(error_msg):
                        error_msg = f"登录SAP失败！Staff ID: {staff_id}, Week: {week}"
                        # logger.error(error_msg)
                        log_data.update({
                            'ID': row['ID'],
                            'staff_id': current_staff_id,
                            'week': current_week,
                            'order_no': row['order_no'],
                            'allocated_hours': row['allocated_hours'],
                            'office_time': row['office_time'],
                            'material_code': row['material_code'],
                            'item': row['item'],
                            'allocated_day': row['allocated_day'],
                            'staff_name': row['staff_name'],
                            'status': 'Failed',
                            'message': error_msg
                        })


                    current_staff_id = staff_id
                    current_week = week
                    is_first_login = False

                # 记录工时
                try:
                    # 准备工时数据
                    # hour_data = {
                    #     'staff_id': staff_id,
                    #     'week': week,
                    #     'order_no': row['order_no'],
                    #     'hours': row['hours'],
                    #     'department': row['department'],
                    #     'project': row['project'],
                    #     'description': row['description']
                    # }
                    hour_data = row
                    # 调用recording_hours方法记录工时
                    recording_res = sap.recording_hours(hour_data)
                    if not recording_res['flag']:
                        # logger.error(error_msg):
                        error_msg = f"记录工时失败！Staff ID: {staff_id}, Week: {week}"
                        # logger.error(error_msg)
                        log_data.update({
                            'ID': row['ID'],
                            'staff_id': current_staff_id,
                            'week': current_week,
                            'order_no': row['order_no'],
                            'allocated_hours': row['allocated_hours'],
                            'office_time': row['office_time'],
                            'material_code': row['material_code'],
                            'item': row['item'],
                            'allocated_day': row['allocated_day'],
                            'staff_name': row['staff_name'],
                            'status': 'Failed',
                            'message': error_msg
                        })


                    success_msg = f"成功处理 Staff ID: {staff_id}, Week: {week} 的工时数据"
                    # logger.info(success_msg)
                    log_data.update({
                        'ID': row['ID'],
                        'staff_id': current_staff_id,
                        'week': current_week,
                        'order_no': row['order_no'],
                        'allocated_hours': row['allocated_hours'],
                        'office_time': row['office_time'],
                        'material_code': row['material_code'],
                        'item': row['item'],
                        'allocated_day': row['allocated_day'],
                        'staff_name': row['staff_name'],
                        'status': 'Success',
                        'message': success_msg
                    })

                    self.textBrowser_4.append(success_msg)
                    QApplication.processEvents()

                except Exception as e:
                    error_msg = f"处理工时数据时出错: {str(e)}"
                    # logger.error(error_msg)
                    log_data.update({
                        'ID': row['ID'],
                        'staff_id': current_staff_id,
                        'week': current_week,
                        'order_no': row['order_no'],
                        'allocated_hours': row['allocated_hours'],
                        'office_time': row['office_time'],
                        'material_code': row['material_code'],
                        'item': row['item'],
                        'allocated_day': row['allocated_day'],
                        'staff_name': row['staff_name'],
                        'status': 'Failed',
                        'message': error_msg
                    })


                log_obj.log(log_data)

            # 最后一次保存
            if not is_first_login:
                save_res = sap.save_hours()
                if not save_res['flag']:
                    error_msg = f"最后一次保存工时失败！Staff ID: {current_staff_id}, Week: {current_week}"
                    # logger.error(error_msg)
                    log_data.update({
                        'ID': row['ID'],
                        'staff_id': current_staff_id,
                        'week': current_week,
                        'order_no': row['order_no'],
                        'allocated_hours': row['allocated_hours'],
                        'office_time': row['office_time'],
                        'material_code': row['material_code'],
                        'item': row['item'],
                        'allocated_day': row['allocated_day'],
                        'staff_name': row['staff_name'],
                        'status': 'Failed',
                        'message': error_msg
                    })


            # 将日志数据保存为Excel文件
            # log_df = pd.DataFrame(log_data)
            # log_file_path = os.path.join(os.path.dirname(hour_path),
            #                              f'hour_operation_log_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
            # log_df.to_excel(log_file_path, index=False)
            log_obj.save_log_to_excel()
            self.textBrowser_4.append("完成", f"所有工时数据处理完成！\n日志文件保存在：{log_file}")
            os.startfile(log_file)
            QApplication.processEvents()

        except Exception as e:
            log_obj.save_log_to_excel()
            os.startfile(log_file)
            self.textBrowser_4.append(f"错误：处理过程中出现错误: {str(e)}\n日志文件保存在：{log_file}")


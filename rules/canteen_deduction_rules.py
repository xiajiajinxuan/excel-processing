import pandas as pd
import numpy as np
from datetime import datetime

def process(data_df, **kwargs):
    """
    处理单个Excel文件中的消费记录和打卡记录，生成扣缴记录和月度汇总
    
    参数:
        data_df (pandas.DataFrame): 此参数在本规则中不使用
        kwargs: 其他参数，包括excel_file（Excel文件路径）
        
    返回:
        dict: 包含扣缴记录和月度汇总的字典
    """
    # 获取Excel文件路径
    excel_file = kwargs.get('excel_file')
    if excel_file is None:
        return {
            'error': '未提供Excel文件路径',
            'deduction_record': pd.DataFrame(),
            'monthly_summary': pd.DataFrame()
        }
    
    try:
        # 读取Excel文件中的工作表
        excel_file_obj = pd.ExcelFile(excel_file)
        sheet_names = excel_file_obj.sheet_names
        
        # 检查是否存在必要的工作表
        if "消费记录" not in sheet_names or "打卡记录" not in sheet_names:
            return {
                'error': 'Excel文件必须包含"消费记录"和"打卡记录"工作表',
                'deduction_record': pd.DataFrame(),
                'monthly_summary': pd.DataFrame()
            }
        
        # 读取工作表数据
        consumption_df = pd.read_excel(excel_file, sheet_name="消费记录")
        attendance_df = pd.read_excel(excel_file, sheet_name="打卡记录")
        
    except Exception as e:
        return {
            'error': f'读取Excel文件时出错: {str(e)}',
            'deduction_record': pd.DataFrame(),
            'monthly_summary': pd.DataFrame()
        }
    
    # 检查打卡记录是否为空
    if attendance_df.empty:
        return {
            'error': '打卡记录为空',
            'deduction_record': pd.DataFrame(),
            'monthly_summary': pd.DataFrame()
        }
    
    # 检查必要的列是否存在于打卡记录
    required_columns_attendance = ['姓名', '工号', '日期', '实际出勤工时']
    missing_columns_attendance = [col for col in required_columns_attendance if col not in attendance_df.columns]
    
    if missing_columns_attendance:
        return {
            'error': f'打卡记录缺少必要的列: {", ".join(missing_columns_attendance)}',
            'deduction_record': pd.DataFrame(),
            'monthly_summary': pd.DataFrame()
        }
    
    # 检查必要的列是否存在于消费记录
    required_columns_consumption = ['工号', '姓名', '交易金额', '交易日期', '餐别']
    if not consumption_df.empty:
        missing_columns_consumption = [col for col in required_columns_consumption if col not in consumption_df.columns]
        
        if missing_columns_consumption:
            return {
                'error': f'消费记录缺少必要的列: {", ".join(missing_columns_consumption)}',
                'deduction_record': pd.DataFrame(),
                'monthly_summary': pd.DataFrame()
            }
    
    # 创建数据副本，避免修改原始数据
    attendance_df = attendance_df.copy()
    consumption_df = consumption_df.copy()
    
    # 确保日期列的格式正确
    try:
        # 转换打卡记录的日期列
        if attendance_df['日期'].dtype == 'object':
            attendance_df['日期'] = pd.to_datetime(attendance_df['日期'], errors='coerce')
        
        # 转换消费记录的日期列
        if not consumption_df.empty and consumption_df['交易日期'].dtype == 'object':
            consumption_df['交易日期'] = pd.to_datetime(consumption_df['交易日期'], errors='coerce')
        
        # 检查是否有无效日期
        if attendance_df['日期'].isna().any():
            # 删除无效日期的行
            attendance_df = attendance_df.dropna(subset=['日期'])
        
        if not consumption_df.empty and consumption_df['交易日期'].isna().any():
            # 删除无效日期的行
            consumption_df = consumption_df.dropna(subset=['交易日期'])
    except Exception as e:
        return {
            'error': f'日期格式转换错误: {str(e)}',
            'deduction_record': pd.DataFrame(),
            'monthly_summary': pd.DataFrame()
        }
    
    # 创建年月列
    attendance_df['年月'] = attendance_df['日期'].dt.strftime('%Y-%m')
    
    # 创建结果DataFrame
    result_df = attendance_df[['姓名', '工号', '日期', '年月', '实际出勤工时']].copy()
    
    # 1. 计算享受就餐减免次数
    def calculate_meal_exemption(hours):
        if hours < 4:
            return 0
        elif hours < 8:
            return 1
        else:
            return 2
    
    result_df['享受就餐减免次数'] = result_df['实际出勤工时'].apply(calculate_meal_exemption)
    
    # 2. 计算实际就餐次数
    result_df['实际就餐次数'] = 0
    result_df['实际就餐金额'] = 0.0
    result_df['消费明细'] = ''
    
    # 处理每一行打卡记录
    for idx, row in result_df.iterrows():
        # 查找对应的消费记录
        if not consumption_df.empty:
            matching_records = consumption_df[
                (consumption_df['工号'] == row['工号']) & 
                (consumption_df['交易日期'].dt.date == row['日期'].date())
            ]
            
            # 计算实际就餐次数
            meal_count = len(matching_records)
            result_df.at[idx, '实际就餐次数'] = meal_count
            
            # 计算实际就餐金额和消费明细
            if meal_count > 0:
                total_amount = 0.0
                meal_details = []
                
                for _, meal_row in matching_records.iterrows():
                    meal_type = meal_row['餐别']
                    # 根据餐别计算金额
                    if meal_type == '早餐':
                        amount = 4.3
                    else:
                        amount = 7.0
                    
                    total_amount += amount
                    meal_details.append(f"{meal_type}:{amount}")
                
                result_df.at[idx, '实际就餐金额'] = round(total_amount, 2)
                result_df.at[idx, '消费明细'] = ';'.join(meal_details)
    
    # 4. 计算应扣就餐减免次数
    result_df['应扣就餐减免次数'] = result_df['实际就餐次数'] - result_df['享受就餐减免次数']
    result_df['应扣就餐减免次数'] = result_df['应扣就餐减免次数'].apply(lambda x: max(0, x))
    
    # 5. 计算应扣金额
    def calculate_deduction_amount(row):
        if row['实际就餐次数'] == 0:
            return 0
        else:
            avg_meal_cost = row['实际就餐金额'] / row['实际就餐次数']
            return round(avg_meal_cost * row['应扣就餐减免次数'], 2)
    
    result_df['应扣金额'] = result_df.apply(calculate_deduction_amount, axis=1)
    
    # 添加备注列
    result_df['备注'] = ''
    
    # 创建月度汇总
    monthly_summary = result_df.groupby(['工号', '年月']).agg({
        '姓名': lambda x: x.iloc[-1],  # 取最后一个姓名
        '享受就餐减免次数': 'sum',
        '实际就餐次数': 'sum',
        '实际就餐金额': 'sum',
        '应扣就餐减免次数': 'sum',
        '应扣金额': 'sum'
    }).reset_index()
    
    # 调整列顺序
    monthly_summary = monthly_summary[['姓名', '工号', '年月', '享受就餐减免次数', '实际就餐次数', 
                                      '实际就餐金额', '应扣就餐减免次数', '应扣金额']]
    
    # 四舍五入金额列到2位小数
    result_df['实际就餐金额'] = result_df['实际就餐金额'].round(2)
    result_df['应扣金额'] = result_df['应扣金额'].round(2)
    monthly_summary['实际就餐金额'] = monthly_summary['实际就餐金额'].round(2)
    monthly_summary['应扣金额'] = monthly_summary['应扣金额'].round(2)
    
    # 返回扣缴记录和月度汇总
    return {
        'deduction_record': result_df,
        'monthly_summary': monthly_summary
    }

def get_rule_info():
    """返回规则的描述信息"""
    return {
        "name": "单Excel文件餐饮消费扣缴规则",
        "description": "处理单个Excel文件中的消费记录和打卡记录工作表，计算员工的就餐减免次数、实际就餐次数、实际就餐金额、应扣就餐减免次数和应扣金额，并生成扣缴记录和月度汇总工作表",
        "version": "1.0",
        "author": "系统"
    } 
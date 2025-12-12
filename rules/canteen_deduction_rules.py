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
    
    # 1. 计算享受就餐减免次数 - 使用向量化操作优化
    # 使用numpy的where函数进行向量化计算，比apply快得多
    result_df['享受就餐减免次数'] = np.where(
        result_df['实际出勤工时'] < 4, 0,
        np.where(result_df['实际出勤工时'] < 8, 1, 2)
    )
    
    # 2. 计算实际就餐次数和金额 - 使用merge优化，避免循环
    # 初始化列
    result_df['实际就餐次数'] = 0
    result_df['实际就餐金额'] = 0.0
    result_df['消费明细'] = ''
    
    if not consumption_df.empty:
        # 为消费记录添加日期列（仅日期部分，用于匹配）
        consumption_df['交易日期_日期'] = consumption_df['交易日期'].dt.date
        
        # 为打卡记录添加日期列（仅日期部分，用于匹配）
        result_df['日期_日期'] = result_df['日期'].dt.date
        
        # 优化：直接对消费记录进行分组统计，然后merge回result_df
        # 计算每餐的金额（向量化操作）
        consumption_df['餐别金额'] = np.where(consumption_df['餐别'] == '早餐', 4.3, 7.0)
        
        # 计算消费明细字符串
        consumption_df['消费明细项'] = consumption_df['餐别'].astype(str) + ':' + consumption_df['餐别金额'].astype(str)
        
        # 按工号和日期分组，计算就餐次数和金额
        meal_stats = consumption_df.groupby(['工号', '交易日期_日期']).agg({
            '餐别': 'count',  # 就餐次数
            '餐别金额': 'sum',  # 总金额
            '消费明细项': lambda x: ';'.join(x.astype(str))  # 消费明细
        }).reset_index()
        
        # 重命名列，统一使用'日期_日期'作为列名以便merge，并重命名统计列
        meal_stats = meal_stats.rename(columns={
            '交易日期_日期': '日期_日期',
            '餐别': '实际就餐次数',
            '餐别金额': '实际就餐金额',
            '消费明细项': '消费明细'
        })
        
        # 将统计结果合并回result_df
        # 先删除已初始化的列，避免merge时的列名冲突
        result_df_temp = result_df.drop(columns=['实际就餐次数', '实际就餐金额', '消费明细'])
        result_df = pd.merge(
            result_df_temp,
            meal_stats,
            on=['工号', '日期_日期'],
            how='left'
        )
        
        # 填充缺失值（没有匹配记录的情况）
        result_df['实际就餐次数'] = result_df['实际就餐次数'].fillna(0).astype(int)
        result_df['实际就餐金额'] = result_df['实际就餐金额'].fillna(0.0).round(2)
        result_df['消费明细'] = result_df['消费明细'].fillna('')
        
        # 删除临时列
        if '日期_日期' in result_df.columns:
            result_df = result_df.drop(columns=['日期_日期'])
    
    # 4. 计算应扣就餐减免次数 - 使用向量化操作
    result_df['应扣就餐减免次数'] = (result_df['实际就餐次数'] - result_df['享受就餐减免次数']).clip(lower=0)
    
    # 5. 计算应扣金额 - 使用向量化操作代替apply
    # 使用numpy的where进行条件计算
    result_df['应扣金额'] = np.where(
        result_df['实际就餐次数'] == 0,
        0,
        (result_df['实际就餐金额'] / result_df['实际就餐次数'] * result_df['应扣就餐减免次数']).round(2)
    )
    
    # 添加备注列
    result_df['备注'] = ''
    
    # 创建月度汇总 - 优化聚合函数
    monthly_summary = result_df.groupby(['工号', '年月']).agg({
        '姓名': 'last',  # 使用'last'代替lambda函数，更快
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
        "version": "1.1",
        "author": "系统"
    }

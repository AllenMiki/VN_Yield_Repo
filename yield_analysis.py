# -*- coding: utf-8 -*-
"""
VN-MES良率汇总分析脚本
功能：分析良率数据，匹配产品状态和不良信息，生成最终报告（两个Sheet）
"""

import pandas as pd
import json
from datetime import datetime
import os

# 设置基础路径
BASE_PATH = r"c:\Users\W13005922\Desktop\资料整理\软件\VN-MES良率汇总\1001-yield"

def load_config():
    """加载配置文件"""
    config_path = os.path.join(BASE_PATH, "config.json")
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_excel_files():
    """加载三个Excel文件"""
    yield_file = os.path.join(BASE_PATH, "良率report_2125_20261001075714.xlsx")
    product_file = os.path.join(BASE_PATH, "产品状态查询report_2217_20261001085927.xlsx")
    defect_file = os.path.join(BASE_PATH, "不良信息汇总report_2113_20261001075911.xlsx")
    
    yield_df = pd.read_excel(yield_file)
    product_df = pd.read_excel(product_file)
    defect_df = pd.read_excel(defect_file)
    
    return yield_df, product_df, defect_df

def filter_yield_data(yield_df, config):
    """
    从良率report中筛选数据
    条件：工站包含关注站位 AND 总不良 > 0
    """
    target_stations = config.get("关注站位", [])
    use_inclusive_match = config.get("匹配规则", {}).get("包含匹配", True)
    
    if use_inclusive_match:
        # 使用包含匹配
        mask = yield_df['总不良'] > 0
        for station in target_stations:
            mask = mask & yield_df['工站'].str.contains(station, na=False)
        filtered_df = yield_df[mask].copy()
    else:
        # 使用精确匹配
        mask = (yield_df['总不良'] > 0) & (yield_df['工站'].isin(target_stations))
        filtered_df = yield_df[mask].copy()
    
    return filtered_df

def match_product_status(filtered_df, product_df):
    """
    根据筛选出的工站名，在产品状态report中查找当前工站相同的记录
    返回匹配的SN列表
    """
    matched_records = []
    
    for _, row in filtered_df.iterrows():
        station = row['工站']
        # 在产品状态中查找当前工站匹配的记录
        matched_products = product_df[product_df['当前工站'] == station]
        
        for _, prod_row in matched_products.iterrows():
            matched_records.append({
                '料号': row['料号'],
                '生产线别': row['生产线别'],
                '工站': station,
                '总不良数': row['总不良'],
                '产品条码': prod_row['产品条码']
            })
    
    return pd.DataFrame(matched_records)

def match_defect_info(matched_products_df, defect_df):
    """
    用产品条码在不良信息report中匹配SERIAL_NUMBER
    按CREATE_TIME取最新的一条记录
    """
    result_records = []
    
    # 按产品条码分组处理
    for product_barcode in matched_products_df['产品条码'].unique():
        # 获取该产品的基本信息和不良信息
        product_info = matched_products_df[matched_products_df['产品条码'] == product_barcode].iloc[0]
        defect_info = defect_df[defect_df['SERIAL_NUMBER'] == product_barcode]
        
        if len(defect_info) > 0:
            # 按CREATE_TIME排序，取最新的一条
            defect_info = defect_info.sort_values('CREATE_TIME', ascending=False).iloc[0]
            
            result_records.append({
                '料号': product_info['料号'],
                '生产线别': product_info['生产线别'],
                '工站': product_info['工站'],
                '总不良数': product_info['总不良数'],
                'SN/产品条码': product_barcode,
                '不良时间': defect_info['CREATE_TIME'],
                '不良工站': defect_info['STATION_NAME'],
                '不良代码': defect_info['DEFECT_CODE'],
                '不良描述': defect_info['DEFECT_DESC'],
                'Machine': defect_info['MACHINE_CODE']
            })
    
    return pd.DataFrame(result_records)

def generate_report(filtered_df, result_df):
    """生成最终Excel报告（两个Sheet）"""
    # 生成文件名，包含时间戳
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_file = os.path.join(BASE_PATH, f"最终良率分析报告_{timestamp}.xlsx")
    
    # Sheet2列顺序
    sn_columns_order = [
        '料号', '生产线别', '工站', '总不良数', 'SN/产品条码',
        '不良时间', '不良工站', '不良代码', '不良描述', 'Machine'
    ]
    
    # 确保Sheet2列顺序
    if len(result_df) > 0:
        result_df = result_df[sn_columns_order]
    
    # 保存到Excel（两个Sheet）
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet1: 良率汇总 - 直接复制原始数据（包含所有16列）
        filtered_df.to_excel(writer, sheet_name='良率汇总', index=False, engine='openpyxl')
        
        # Sheet2: SN详细信息
        result_df.to_excel(writer, sheet_name='SN详细信息', index=False, engine='openpyxl')
    
    return output_file

def main():
    """主函数"""
    print("=" * 60)
    print("VN-MES 良率汇总分析程序")
    print("=" * 60)
    
    # 1. 加载配置
    print("\n[1] 加载配置文件...")
    config = load_config()
    print(f"    关注站位: {config.get('关注站位')}")
    print(f"    匹配规则: {config.get('匹配规则')}")
    
    # 2. 加载Excel文件
    print("\n[2] 加载Excel文件...")
    yield_df, product_df, defect_df = load_excel_files()
    print(f"    良率report: {len(yield_df)} 行")
    print(f"    产品状态report: {len(product_df)} 行")
    print(f"    不良信息report: {len(defect_df)} 行")
    
    # 3. 筛选良率数据
    print("\n[3] 筛选良率数据...")
    filtered_df = filter_yield_data(yield_df, config)
    station_list = filtered_df['工站'].unique().tolist()
    print(f"    筛选出的工站列表: {station_list}")
    print(f"    筛选出的记录数: {len(filtered_df)}")
    
    # 4. 匹配产品状态
    print("\n[4] 匹配产品状态...")
    matched_products_df = match_product_status(filtered_df, product_df)
    print(f"    匹配到的产品数量: {len(matched_products_df['产品条码'].unique())}")
    
    # 5. 匹配不良信息
    print("\n[5] 匹配不良信息...")
    result_df = match_defect_info(matched_products_df, defect_df)
    print(f"    最终报告行数: {len(result_df)}")
    
    # 6. 生成报告（两个Sheet）
    print("\n[6] 生成最终报告（两个Sheet）...")
    output_file = generate_report(filtered_df, result_df)
    print(f"    报告已保存至: {output_file}")
    
    # 显示Sheet1预览
    if len(filtered_df) > 0:
        print("\n[7] Sheet1-良率汇总 预览 (前5行):")
        print(filtered_df.head().to_string())
    
    # 显示Sheet2预览
    if len(result_df) > 0:
        print("\n[8] Sheet2-SN详细信息 预览 (前5行):")
        print(result_df.head().to_string())
    
    print("\n" + "=" * 60)
    print("分析完成!")
    print("=" * 60)
    
    return filtered_df, result_df

if __name__ == "__main__":
    filtered_result, sn_result = main()
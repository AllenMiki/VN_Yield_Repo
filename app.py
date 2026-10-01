# -*- coding: utf-8 -*-
"""
VN-MES良率汇总分析工具 - Streamlit Web应用
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(
    page_title="VN-MES良率汇总分析工具",
    page_icon="📊",
    layout="centered"
)

# 默认配置
DEFAULT_STATIONS = "初测"


def load_excel_files(yield_file, product_file, defect_file):
    yield_df = pd.read_excel(yield_file)
    product_df = pd.read_excel(product_file)
    defect_df = pd.read_excel(defect_file)
    return yield_df, product_df, defect_df


def filter_yield_data(yield_df, target_stations, use_inclusive_match=True):
    if not target_stations:
        return pd.DataFrame()
    
    if use_inclusive_match:
        mask = yield_df['总不良'] > 0
        for station in target_stations:
            mask = mask & yield_df['工站'].str.contains(station, na=False)
        return yield_df[mask].copy()
    else:
        mask = (yield_df['总不良'] > 0) & (yield_df['工站'].isin(target_stations))
        return yield_df[mask].copy()


def match_product_status(filtered_df, product_df):
    matched_records = []
    for _, row in filtered_df.iterrows():
        station = row['工站']
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
    result_records = []
    if len(matched_products_df) == 0:
        return pd.DataFrame()
    
    for product_barcode in matched_products_df['产品条码'].unique():
        product_info = matched_products_df[matched_products_df['产品条码'] == product_barcode].iloc[0]
        defect_info = defect_df[defect_df['SERIAL_NUMBER'] == product_barcode]
        
        if len(defect_info) > 0:
            defect_info = defect_info.sort_values('CREATE_TIME', ascending=False).iloc[0]
            result_records.append({
                '料号': product_info['料号'],
                '生产线别': product_info['生产线别'],
                '工站': product_info['工站'],
                'SN/产品条码': product_barcode,
                '不良时间': defect_info['CREATE_TIME'],
                '不良工站': defect_info['STATION_NAME'],
                '不良代码': defect_info['DEFECT_CODE'],
                '不良描述': defect_info['DEFECT_DESC'],
                'Machine': defect_info['MACHINE_CODE']
            })
    return pd.DataFrame(result_records)


def generate_report(filtered_df, result_df):
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_filename = f"最终良率分析报告_{timestamp}.xlsx"
    
    sn_columns_order = ['料号', '生产线别', '工站', 'SN/产品条码',
                        '不良时间', '不良工站', '不良代码', '不良描述', 'Machine']
    result_df_copy = result_df.copy()
    if len(result_df_copy) > 0:
        result_df_copy = result_df_copy[sn_columns_order]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, sheet_name='良率汇总', index=False)
        result_df_copy.to_excel(writer, sheet_name='SN详细信息', index=False)
    
    output.seek(0)
    return output, output_filename


# 界面
st.markdown("# VN-MES 良率汇总分析工具")
st.markdown("---")

# 侧边栏
with st.sidebar:
    st.markdown("## 设置")
    
    yield_file = st.file_uploader("良率 report (xlsx)", type=['xlsx'])
    product_file = st.file_uploader("产品状态 report (xlsx)", type=['xlsx'])
    defect_file = st.file_uploader("不良信息 report (xlsx)", type=['xlsx'])
    
    st.markdown("---")
    st.markdown("### 关注站位配置")
    
    stations_input = st.text_input(
        "输入站位（用逗号分隔）",
        value=DEFAULT_STATIONS,
        help="例如: 初测,复测"
    )
    
    use_inclusive_match = st.checkbox("包含匹配", value=True)
    match_rule = "包含匹配" if use_inclusive_match else "精确匹配"
    
    st.markdown(f"**匹配规则:** {match_rule}")
    
    st.markdown("---")
    analyze_button = st.button("🚀 开始分析", type="primary", use_container_width=True)
    clear_button = st.button("🗑️ 清空", use_container_width=True)

# 主界面
if analyze_button:
    # 解析站位
    target_stations = [s.strip() for s in stations_input.split(',') if s.strip()]
    
    if not target_stations:
        st.error("请配置至少一个关注站位！")
    elif not yield_file or not product_file or not defect_file:
        st.error("请上传全部三个Excel文件！")
    else:
        try:
            with st.spinner("正在处理..."):
                # 加载文件
                yield_df, product_df, defect_df = load_excel_files(yield_file, product_file, defect_file)
                
                # 筛选
                filtered_df = filter_yield_data(yield_df, target_stations, use_inclusive_match)
                matched_products_df = match_product_status(filtered_df, product_df)
                result_df = match_defect_info(matched_products_df, defect_df)
                report_data, report_filename = generate_report(filtered_df, result_df)
            
            # 显示统计
            st.markdown("### 统计信息")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("关注站位", len(target_stations))
            with col2:
                st.metric("筛选记录", len(filtered_df))
            with col3:
                st.metric("匹配产品", len(matched_products_df['产品条码'].unique()) if len(matched_products_df) > 0 else 0)
            with col4:
                st.metric("最终记录", len(result_df))
            
            # Sheet1
            st.markdown("### Sheet1: 良率汇总")
            if len(filtered_df) > 0:
                st.dataframe(filtered_df, use_container_width=True, height=300)
            else:
                st.info("无符合条件的数据")
            
            st.markdown("---")
            
            # Sheet2
            st.markdown("### Sheet2: SN详细信息")
            if len(result_df) > 0:
                sn_columns = ['料号', '生产线别', '工站', 'SN/产品条码',
                              '不良时间', '不良工站', '不良代码', '不良描述', 'Machine']
                st.dataframe(result_df[sn_columns], use_container_width=True, height=300)
            else:
                st.info("无匹配的不良信息")
            
            st.markdown("---")
            
            # 下载
            st.markdown("### 导出报告")
            st.download_button(
                label="📥 下载Excel报告",
                data=report_data,
                file_name=report_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.success("✅ 分析完成！")
            
        except Exception as e:
            st.error(f"错误: {str(e)}")

if clear_button:
    st.rerun()

# 初始界面
if 'filtered_df' not in dir() or (analyze_button and len(filtered_df) == 0 and not st.session_state.get('processed')):
    st.markdown("""
    ### 使用说明
    1. 在左侧上传三个Excel文件（良率、产品状态、不良信息）
    2. 配置关注站位（用逗号分隔，如：初测,复测）
    3. 选择匹配规则（包含匹配或精确匹配）
    4. 点击「开始分析」按钮
    """)
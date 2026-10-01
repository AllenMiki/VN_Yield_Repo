# -*- coding: utf-8 -*-
"""
VN-MES良率汇总分析工具 - Streamlit Web应用
功能：分析良率数据，匹配产品状态和不良信息，生成最终报告
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 设置页面配置
st.set_page_config(
    page_title="VN-MES良率汇总分析工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 默认配置（用于云端部署）
DEFAULT_CONFIG = {"关注站位": ["初测"], "匹配规则": {"包含匹配": True}}


def load_excel_files(yield_file, product_file, defect_file):
    """
    从上传的文件对象加载Excel
    """
    yield_df = pd.read_excel(yield_file)
    product_df = pd.read_excel(product_file)
    defect_df = pd.read_excel(defect_file)
    return yield_df, product_df, defect_df


def filter_yield_data(yield_df, target_stations, use_inclusive_match=True):
    """
    从良率report中筛选数据
    条件：工站包含关注站位 AND 总不良 > 0
    """
    if not target_stations:
        return pd.DataFrame()
    
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
    
    if len(matched_products_df) == 0:
        return pd.DataFrame()
    
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
                'SN/产品条码': product_barcode,
                '不良时间': defect_info['CREATE_TIME'],
                '不良工站': defect_info['STATION_NAME'],
                '不良代码': defect_info['DEFECT_CODE'],
                '不良描述': defect_info['DEFECT_DESC'],
                'Machine': defect_info['MACHINE_CODE']
            })
    
    return pd.DataFrame(result_records)


def generate_report(filtered_df, result_df):
    """
    生成最终Excel报告（两个Sheet）
    返回内存中的Excel文件数据
    """
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_filename = f"最终良率分析报告_{timestamp}.xlsx"
    
    # Sheet2列顺序
    sn_columns_order = [
        '料号', '生产线别', '工站', 'SN/产品条码',
        '不良时间', '不良工站', '不良代码', '不良描述', 'Machine'
    ]
    
    # 复制数据避免修改原数据
    result_df_copy = result_df.copy()
    
    # 确保Sheet2列顺序
    if len(result_df_copy) > 0:
        result_df_copy = result_df_copy[sn_columns_order]
    
    # 在内存中生成Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet1: 良率汇总
        filtered_df.to_excel(writer, sheet_name='良率汇总', index=False, engine='openpyxl')
        
        # Sheet2: SN详细信息
        result_df_copy.to_excel(writer, sheet_name='SN详细信息', index=False, engine='openpyxl')
    
    output.seek(0)
    return output, output_filename


# ==================== Streamlit 界面 ====================

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-size: 1.1rem;
        padding: 0.75rem 2rem;
    }
    .stButton>button:hover {
        background-color: #1666a5;
    }
    .info-box {
        background-color: #e7f3ff;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #2196F3;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
        margin-bottom: 1rem;
    }
    .error-box {
        background-color: #f8d7da;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #dc3545;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
        margin-bottom: 1rem;
    }
    /* 文件上传区域样式 */
    .file-uploader-section {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    /* 配置编辑区域样式 */
    .config-section {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# 主标题
st.markdown('<p class="main-header">VN-MES 良率汇总分析工具</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">高效分析良率数据，匹配产品状态与不良信息</p>', unsafe_allow_html=True)

# ==================== 初始化会话状态 ====================
if 'config' not in st.session_state:
    st.session_state['config'] = DEFAULT_CONFIG.copy()

if 'stations_list' not in st.session_state:
    # 初始化站位列表为DataFrame格式，便于编辑
    stations = st.session_state['config'].get('关注站位', ["初测"])
    st.session_state['stations_list'] = pd.DataFrame({"关注站位": stations})

if 'use_inclusive_match' not in st.session_state:
    match_rule = st.session_state['config'].get('匹配规则', {})
    st.session_state['use_inclusive_match'] = match_rule.get('包含匹配', True)

# ==================== 侧边栏设置 ====================
with st.sidebar:
    st.markdown("## ⚙️ 设置面板")
    st.markdown("---")
    
    # 文件上传区域
    st.markdown("### 📁 数据文件上传")
    
    st.markdown('<div class="file-uploader-section">', unsafe_allow_html=True)
    
    yield_file = st.file_uploader(
        "良率 report (xlsx)",
        type=['xlsx'],
        help="上传良率report Excel文件"
    )
    
    product_file = st.file_uploader(
        "产品状态 report (xlsx)",
        type=['xlsx'],
        help="上传产品状态查询report Excel文件"
    )
    
    defect_file = st.file_uploader(
        "不良信息 report (xlsx)",
        type=['xlsx'],
        help="上传不良信息汇总report Excel文件"
    )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 关注站位编辑区域
    st.markdown("### 🔧 关注站位配置")
    
    st.markdown('<div class="config-section">', unsafe_allow_html=True)
    
    # 使用 data_editor 编辑站位列表
    st.markdown("**编辑站位列表（可直接添加/删除/修改）：**")
    
    edited_df = st.data_editor(
        st.session_state['stations_list'],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="stations_editor"
    )
    
    # 保存配置按钮
    col_save1, col_save2 = st.columns(2)
    
    with col_save1:
        if st.button("💾 确认站位配置", use_container_width=True):
            # 从编辑后的DataFrame提取站位列表
            stations_list = edited_df['关注站位'].dropna().tolist()
            stations_list = [s.strip() for s in stations_list if str(s).strip()]
            
            st.session_state['config'] = {
                "关注站位": stations_list,
                "匹配规则": {"包含匹配": st.session_state['use_inclusive_match']}
            }
            st.session_state['stations_list'] = edited_df.copy()
            st.success("✅ 站位配置已更新（仅当前会话有效）")
    
    with col_save2:
        if st.button("🔄 重置为默认配置", use_container_width=True):
            st.session_state['config'] = DEFAULT_CONFIG.copy()
            st.session_state['stations_list'] = pd.DataFrame({"关注站位": ["初测"]})
            st.session_state['use_inclusive_match'] = True
            st.rerun()
    
    st.caption("💡 注意：云端部署无法保存配置到文件，刷新页面将重置")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 筛选条件
    st.markdown("### 🔍 筛选条件")
    
    use_inclusive_match = st.checkbox(
        "包含匹配",
        value=st.session_state['use_inclusive_match'],
        help="开启：工站名称包含输入的站位即匹配；关闭：工站名称需精确匹配"
    )
    st.session_state['use_inclusive_match'] = use_inclusive_match
    
    match_rule_text = "包含匹配" if use_inclusive_match else "精确匹配"
    
    # 显示当前配置的站位
    current_stations = st.session_state['stations_list']['关注站位'].dropna().tolist()
    current_stations = [s for s in current_stations if str(s).strip()]
    if current_stations:
        st.markdown(f"**当前站位:** {', '.join(current_stations)}")
    else:
        st.markdown("**当前站位:** <无配置>")
    
    st.markdown("---")
    
    # 分析按钮
    analyze_button = st.button("🚀 开始分析", type="primary", use_container_width=True)
    
    # 清空按钮
    clear_button = st.button("🗑️ 清空结果", use_container_width=True)


# ==================== 主界面显示 ====================

# 检查文件是否存在并执行分析
if analyze_button:
    # 解析关注站位
    target_stations = st.session_state['stations_list']['关注站位'].dropna().tolist()
    target_stations = [s.strip() for s in target_stations if str(s).strip()]
    
    if not target_stations:
        st.error("⚠️ 请配置至少一个关注站位！")
    elif yield_file is None or product_file is None or defect_file is None:
        st.error("⚠️ 请上传全部三个Excel文件！")
    else:
        # 进度显示
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # 步骤1: 显示文件信息
            status_text.text("⏳ 正在处理文件...")
            progress_bar.progress(10)
            
            # 步骤2: 加载文件
            status_text.text("⏳ 正在加载Excel文件...")
            progress_bar.progress(20)
            
            yield_df, product_df, defect_df = load_excel_files(
                yield_file, product_file, defect_file
            )
            
            # 在session_state中存储数据
            st.session_state['yield_df'] = yield_df
            st.session_state['product_df'] = product_df
            st.session_state['defect_df'] = defect_df
            
            # 步骤3: 显示文件预览
            status_text.text("⏳ 正在生成文件预览...")
            progress_bar.progress(30)
            
            st.markdown("## 📊 文件预览")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("#### 良率 Report")
                st.markdown(f"**文件名:** {yield_file.name}")
                st.markdown(f"**行数:** {len(yield_df)}")
                st.markdown(f"**列数:** {len(yield_df.columns)}")
                with st.expander("查看列名"):
                    for col in yield_df.columns:
                        st.markdown(f"- {col}")
            
            with col2:
                st.markdown("#### 产品状态 Report")
                st.markdown(f"**文件名:** {product_file.name}")
                st.markdown(f"**行数:** {len(product_df)}")
                st.markdown(f"**列数:** {len(product_df.columns)}")
                with st.expander("查看列名"):
                    for col in product_df.columns:
                        st.markdown(f"- {col}")
            
            with col3:
                st.markdown("#### 不良信息 Report")
                st.markdown(f"**文件名:** {defect_file.name}")
                st.markdown(f"**行数:** {len(defect_df)}")
                st.markdown(f"**列数:** {len(defect_df.columns)}")
                with st.expander("查看列名"):
                    for col in defect_df.columns:
                        st.markdown(f"- {col}")
            
            st.markdown("---")
            
            # 步骤4: 筛选良率数据
            status_text.text("⏳ 正在筛选良率数据...")
            progress_bar.progress(40)
            
            filtered_df = filter_yield_data(yield_df, target_stations, use_inclusive_match)
            
            # 步骤5: 匹配产品状态
            status_text.text("⏳ 正在匹配产品状态...")
            progress_bar.progress(60)
            
            matched_products_df = match_product_status(filtered_df, product_df)
            
            # 步骤6: 匹配不良信息
            status_text.text("⏳ 正在匹配不良信息...")
            progress_bar.progress(80)
            
            result_df = match_defect_info(matched_products_df, defect_df)
            
            # 步骤7: 生成报告
            status_text.text("⏳ 正在生成报告...")
            progress_bar.progress(90)
            
            report_data, report_filename = generate_report(filtered_df, result_df)
            
            # 在session_state中存储结果
            st.session_state['filtered_df'] = filtered_df
            st.session_state['result_df'] = result_df
            st.session_state['report_data'] = report_data.getvalue()
            st.session_state['report_filename'] = report_filename
            
            progress_bar.progress(100)
            status_text.text("✅ 分析完成！")
            
            # 统计信息
            st.markdown("## 📈 统计信息")
            
            stat_col1, stat_col2, stat_col3, stat_col4, stat_col5 = st.columns(5)
            
            with stat_col1:
                st.metric("关注站位", f"{len(target_stations)}个", ", ".join(target_stations[:3]) + ("..." if len(target_stations) > 3 else ""))
            
            with stat_col2:
                st.metric("匹配规则", match_rule_text)
            
            with stat_col3:
                unique_stations = len(filtered_df['工站'].unique()) if len(filtered_df) > 0 else 0
                st.metric("筛选工站", f"{len(filtered_df)}个", f"符合条件: {unique_stations}个")
            
            with stat_col4:
                unique_products = len(matched_products_df['产品条码'].unique()) if len(matched_products_df) > 0 else 0
                st.metric("匹配产品", f"{unique_products}个")
            
            with stat_col5:
                st.metric("最终记录", f"{len(result_df)}条")
            
            st.markdown("---")
            
            # Sheet1: 良率汇总表格
            st.markdown("## 📋 Sheet1: 良率汇总")
            
            if len(filtered_df) > 0:
                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    height=400
                )
                
                st.markdown(f"**良率汇总表格共 {len(filtered_df)} 条记录**")
            else:
                st.info("ℹ️ 没有符合筛选条件的数据")
            
            st.markdown("---")
            
            # Sheet2: SN详细信息表格
            st.markdown("## 🔍 Sheet2: SN详细信息")
            
            if len(result_df) > 0:
                # SN列顺序
                sn_columns_order = [
                    '料号', '生产线别', '工站', 'SN/产品条码',
                    '不良时间', '不良工站', '不良代码', '不良描述', 'Machine'
                ]
                display_df = result_df[sn_columns_order]
                
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    height=400
                )
                
                st.markdown(f"**SN详细信息表格共 {len(result_df)} 条记录**")
            else:
                st.info("ℹ️ 没有匹配到不良信息记录")
            
            st.markdown("---")
            
            # 下载报告按钮
            st.markdown("## 📥 导出报告")
            
            col_down1, col_down2 = st.columns(2)
            
            with col_down1:
                st.download_button(
                    label="📥 下载完整报告 (Excel)",
                    data=report_data,
                    file_name=report_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col_down2:
                st.markdown(f"📄 文件名: `{report_filename}`")
                st.markdown(f"⏰ 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            st.success("✅ 分析完成！请查看上方报表或下载报告。")
            
        except Exception as e:
            st.error(f"❌ 分析过程中发生错误: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            progress_bar.empty()
            status_text.empty()

# 如果有历史结果，显示它们
elif 'filtered_df' in st.session_state and 'result_df' in st.session_state:
    filtered_df = st.session_state['filtered_df']
    result_df = st.session_state['result_df']
    
    st.markdown("## 📊 历史结果")
    
    if len(filtered_df) > 0:
        st.markdown("### Sheet1: 良率汇总")
        st.dataframe(filtered_df, use_container_width=True, height=400)
    else:
        st.info("ℹ️ 良率汇总无数据")
    
    if len(result_df) > 0:
        st.markdown("### Sheet2: SN详细信息")
        sn_columns_order = [
            '料号', '生产线别', '工站', 'SN/产品条码',
            '不良时间', '不良工站', '不良代码', '不良描述', 'Machine'
        ]
        display_df = result_df[sn_columns_order]
        st.dataframe(display_df, use_container_width=True, height=400)
    else:
        st.info("ℹ️ SN详细信息无数据")
    
    # 显示下载按钮
    if 'report_data' in st.session_state:
        st.markdown("---")
        st.markdown("## 📥 导出报告")
        
        st.download_button(
            label="📥 下载报告 (Excel)",
            data=st.session_state['report_data'],
            file_name=st.session_state['report_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# 清空结果
if clear_button:
    for key in ['yield_df', 'product_df', 'defect_df', 'filtered_df', 'result_df', 'output_file']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

# 初始欢迎界面
if 'filtered_df' not in st.session_state:
    st.markdown("""
    <div class="info-box">
        <h4>👋 欢迎使用VN-MES良率汇总分析工具</h4>
        <p>请在左侧设置面板中完成以下配置：</p>
        <ul>
            <li><strong>上传数据文件</strong> - 依次上传三个Excel报告文件（良率、产品状态、不良信息）</li>
            <li><strong>关注站位配置</strong> - 使用表格编辑器添加/删除/修改站位，点击"确认站位配置"</li>
            <li><strong>匹配规则</strong> - 选择包含匹配或精确匹配</li>
        </ul>
        <p>配置完成后，点击「开始分析」按钮进行分析。</p>
    </div>
    """, unsafe_allow_html=True)
    
    # 显示配置信息
    st.markdown("### 📋 当前配置")
    
    config = st.session_state['config']
    stations = config.get('关注站位', [])
    
    st.markdown("**当前关注站位:**")
    if stations:
        for station in stations:
            st.markdown(f"- {station}")
    else:
        st.markdown("<无配置>")
    
    st.markdown("---")
    
    # 使用说明
    st.markdown("### 📖 使用说明")
    
    usage_col1, usage_col2 = st.columns(2)
    
    with usage_col1:
        st.markdown("""
        **步骤1: 上传文件**
        - 在左侧上传三个 xlsx 格式的 Excel 文件
        - 良率report、产品状态report、不良信息report
        
        **步骤2: 配置站位**
        - 在"关注站位配置"区域编辑站位列表
        - 支持动态添加/删除/修改
        - 点击"确认站位配置"应用更改
        - 点击"重置为默认配置"恢复默认
        """)
    
    with usage_col2:
        st.markdown("""
        **步骤3: 开始分析**
        - 设置匹配规则（包含匹配/精确匹配）
        - 点击"🚀 开始分析"按钮
        
        **步骤4: 查看结果**
        - 查看良率汇总和SN详细信息
        - 下载生成的 Excel 报告
        - 报告包含两个 Sheet:
          - Sheet1: 良率汇总
          - Sheet2: SN详细信息
        """)
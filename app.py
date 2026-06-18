import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime

# ================= 1. 页面配置与高级样式 (Apple Style) =================
st.set_page_config(
    page_title="制衣厂销量看板 Pro",
    layout="wide",
    page_icon="👕",
    initial_sidebar_state="expanded"
)

# 注入 Apple 风格 CSS
st.markdown("""
<style>
    /* 全局字体与背景 */
    .stApp {
        background-color: #F5F5F7; /* Apple 浅灰背景 */
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    /* 标题样式 */
    h1 { color: #1D1D1F; font-weight: 700; letter-spacing: -0.02em; margin-bottom: 20px; }
    h2, h3 { color: #1D1D1F; font-weight: 600; }

    /* 卡片容器样式 (模拟 iOS 小组件) */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border: 1px solid rgba(0,0,0,0.02);
    }

    /* 侧边栏样式 */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E5E5EA;
    }
    section[data-testid="stSidebar"] h1 { font-size: 20px; margin-top: 10px;}

    /* 按钮与输入框圆角 */
    .stButton>button {
        border-radius: 20px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stDateInput > div > div > input {
        border-radius: 12px;
    }

    /* 图表容器 */
    .chart-container {
        background-color: #FFFFFF;
        padding: 20px;
        border-radius: 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.title("👕 制衣厂销量数据中心")
st.caption("实时生产监控与趋势分析看板")

# ================= 2. 智能数据加载引擎 =================
@st.cache_data(ttl=3600) # 缓存1小时，提升性能
def load_and_clean_data():
    file_path = r"C:\Users\Administrator\Desktop\ClothingDashboard\2025年打吊牌统计表.xlsx"

    if not os.path.exists(file_path):
        st.error(f"未找到文件：{file_path}")
        return pd.DataFrame()

    try:
        # 读取所有Sheet
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        df_list = []

        for sheet_name, df in all_sheets.items():
            # 1. 智能识别列名 (处理空格、别名)
            cols = df.columns.tolist()
            date_col = next((c for c in cols if '日期' in str(c)), None)
            model_col = next((c for c in cols if '型号' in str(c) or '款号' in str(c)), None)
            qty_col = next((c for c in cols if '实际数量' in str(c) or '数量' in str(c)), None)

            if date_col and model_col and qty_col:
                temp_df = df[[date_col, model_col, qty_col]].copy()
                temp_df.columns = ['日期', '型号', '实际数量']
                temp_df['来源分表'] = sheet_name # 标记数据来源
                df_list.append(temp_df)

        if not df_list:
            st.warning("未能识别任何包含【日期】、【型号】、【实际数量】的工作表。")
            return pd.DataFrame()

        # 合并所有数据
        final_df = pd.concat(df_list, ignore_index=True)

        # 2. 强力数据清洗
        # 转换日期 (处理各种可能的格式)
        final_df['日期'] = pd.to_datetime(final_df['日期'], errors='coerce')
        # 转换数量为数字
        final_df['实际数量'] = pd.to_numeric(final_df['实际数量'], errors='coerce').fillna(0)
        # 转换型号为字符串 (解决排序报错问题)
        final_df['型号'] = final_df['型号'].astype(str).str.strip()

        # 删除无效行
        final_df.dropna(subset=['日期'], inplace=True)

        return final_df

    except Exception as e:
        st.error(f"数据读取失败: {e}")
        return pd.DataFrame()

# 执行加载
df = load_and_clean_data()

if df.empty:
    st.stop()

# ================= 3. 侧边栏筛选器 (移动端友好) =================
with st.sidebar:
    st.header("筛选面板")

    # 获取所有唯一型号并排序
    all_models = sorted(df['型号'].unique().tolist())

    # 多选框
    selected_models = st.multiselect(
        "选择对比型号",
        options=all_models,
        default=[all_models[0]] if len(all_models) > 0 else [], # 默认选中第一个
        placeholder="请选择..."
    )

    # 日期范围选择
    min_date = df['日期'].min()
    max_date = df['日期'].max()

    date_range = st.date_input(
        "查询日期范围",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

# ================= 4. 数据处理与过滤 =================
# 处理日期范围选择逻辑
if len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

# 过滤数据
mask = (df['日期'] >= pd.Timestamp(start_date)) & \
       (df['日期'] <= pd.Timestamp(end_date)) & \
       (df['型号'].isin(selected_models))

filtered_df = df[mask].copy()

# ================= 5. 主界面展示 =================
if not selected_models:
    st.info("请在左侧选择一个或多个型号以开始分析。")
    st.stop()

# --- A. 关键指标卡片 (KPI Cards) ---
col1, col2, col3 = st.columns(3)

total_qty = int(filtered_df['实际数量'].sum())
record_count = len(filtered_df)
model_count = len(filtered_df['型号'].unique())

col1.metric(label="总实际数量", value=f"{total_qty:,}", delta=None)
col2.metric(label="涉及型号数", value=model_count)
col3.metric(label="记录条数", value=record_count)

st.markdown("<br>", unsafe_allow_html=True) # 间距

# --- B. 趋势图表 (Plotly) ---
st.subheader("每日产量趋势分析")

if filtered_df.empty:
    st.warning("当前筛选条件下无数据。")
else:
    # 按日期聚合，防止同一天多条记录导致折线图混乱
    daily_trend = filtered_df.groupby(['日期', '型号'])['实际数量'].sum().reset_index()

    fig = px.line(
        daily_trend,
        x='日期',
        y='实际数量',
        color='型号',
        markers=True,
        title="各型号每日实际数量走势",
        template="plotly_white", # 使用干净的白色模板
        color_discrete_sequence=px.colors.qualitative.Set2 # 使用更专业的配色
    )

    # 图表美化
    fig.update_layout(
        font=dict(family="SF Pro Text, sans-serif"),
        plot_bgcolor='rgba(0,0,0,0)',
        hovermode="x unified",
        legend_title_text="型号",
        height=500
    )
    fig.update_xaxes(showgrid=True, gridcolor='#E5E5EA')
    fig.update_yaxes(showgrid=True, gridcolor='#E5E5EA', title_text="实际数量")

    # 在容器中显示图表，增加圆角背景
    with st.container():
        st.plotly_chart(fig, use_container_width=True)

# --- C. 详细数据表 ---
st.markdown("---")
with st.expander("查看详细数据"):
    display_df = filtered_df.sort_values(by='日期', ascending=False)
    # 格式化日期显示
    display_df['日期'] = display_df['日期'].dt.strftime('%Y-%m-%d')
    st.dataframe(display_df, hide_index=True, use_container_width=True)

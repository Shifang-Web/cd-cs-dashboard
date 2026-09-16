import streamlit as st
import pandas as pd
import plotly.express as px
import io

# ==========================================
# 1. 页面基础设置
# ==========================================
st.set_page_config(page_title="项目风险看板", layout="wide")
st.title("🚨 成都城市营业部项目风险看板")

# ==========================================
# 2. 动态数据通道（支持上传、自动合并所有 Sheet）
# ==========================================
st.sidebar.header("📁 数据通道管理")
uploaded_file = st.sidebar.file_uploader(
    "上传新的数据文件（支持 CSV 或 Excel）", 
    type=["csv", "xlsx", "xls"]
)

def load_default_data():
    """没上传文件时，默认读取本地 data.xlsx"""
    try:
        xls = pd.ExcelFile("data.xlsx")
        all_dfs = []
        for sheet in xls.sheet_names:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df['来源工作表'] = sheet
            all_dfs.append(temp_df)
        return pd.concat(all_dfs, ignore_index=True)
    except Exception as e:
        # 如果本地没有 data.xlsx，抛出明确的异常，交给主程序统一捕获处理
        raise RuntimeError(f"本地文件 data.xlsx 读取失败，请检查文件是否存在。原始错误：{e}")

def get_data():
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        if uploaded_file.name.endswith('.csv'):
            try:
                return pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
            except UnicodeDecodeError:
                return pd.read_csv(io.BytesIO(file_bytes), encoding='gbk')
        else:
            # 处理 Excel：自动合并所有 Sheet
            try:
                xls = pd.ExcelFile(io.BytesIO(file_bytes))
                all_dfs = []
                for sheet in xls.sheet_names:
                    temp_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
                    temp_df['来源工作表'] = sheet  # 标记数据来源
                    all_dfs.append(temp_df)
                return pd.concat(all_dfs, ignore_index=True)
            except Exception as e:
                st.sidebar.error(f"⚠️ 读取Excel失败！请确保文件未被其他软件打开，且格式为xlsx。错误：{e}")
                st.stop()
    else:
        return load_default_data()

# ==========================================
# 3. 数据加载（关键修复：预先初始化 df，防止 NameError）
# ==========================================
# 初始化一个空的 DataFrame，确保后续代码不会因为 df 未定义而崩溃
df = pd.DataFrame()

try:
    df = get_data()
    st.sidebar.success(f"数据加载成功！共合并了 {len(df)} 条记录")
except Exception as e:
    st.error(f"读取文件失败，请检查文件名和路径。报错信息：{e}")
    st.stop()

# 如果读取失败并执行了 st.stop()，后续代码会停止（在正确的 Streamlit 环境下）

# ==========================================
# 4. 侧边栏多级筛选
# ==========================================
st.sidebar.header("⚙️ 风险筛选器")

# (0) 来源工作表筛选
if '来源工作表' in df.columns:
    sheet_options = list(df['来源工作表'].dropna().unique())
    selected_sheets = st.sidebar.multiselect(
        "选择来源表格（可多选）：",
        options=sheet_options,
        default=sheet_options
    )
    df = df[df['来源工作表'].isin(selected_sheets)]

# (1) 片区筛选
if '片区' in df.columns:
    filter_options = ['全部'] + list(df['片区'].dropna().unique())
    selected_filter = st.sidebar.selectbox("按片区筛选：", filter_options)
    if selected_filter != '全部':
        df = df[df['片区'] == selected_filter]

# (2) 风险项筛选（注意：请把 '防汛分级' 改成你表格里真实的列名，如 '风险等级'）
risk_col = '防汛分级' 
if risk_col in df.columns:
    risk_options = list(df[risk_col].dropna().unique())
    selected_risks = st.sidebar.multiselect(
        "按风险等级筛选（可多选）：",
        options=risk_options,
        default=risk_options
    )
    df = df[df[risk_col].isin(selected_risks)]

# ==========================================
# 5. 数据可视化（并排双图表 + 指标卡）
# ==========================================
st.subheader("📊 片区 & 蝶城 双维度风险统计")

# 指标卡片
col1, col2, col3 = st.columns(3)
col1.metric("当前筛选下项目总数", len(df))
col2.metric("数据来源表格数", df['来源工作表'].nunique() if '来源工作表' in df.columns else 1)

if risk_col in df.columns:
    high_risk_count = len(df[df[risk_col] == '高风险'])
    col3.metric("🚨 高风险项目数", high_risk_count, delta="需重点关注", delta_color="inverse")

st.markdown("---")

# 并排展示“片区”和“蝶城”
target_dims = [col for col in ['片区', '蝶城'] if col in df.columns]

if not target_dims:
    st.warning("⚠️ 表格中未找到'片区'或'蝶城'列，请检查表头名称是否一致。")
else:
    # 创建两列布局
    chart_cols = st.columns(len(target_dims))
    
    for i, dim in enumerate(target_dims):
        # 统计该维度数据
        chart_data = df[dim].value_counts().reset_index()
        chart_data.columns = ['分类', '数量']
        
        # 画图
        fig = px.bar(
            chart_data, 
            x='分类', 
            y='数量', 
            color='分类', 
            text_auto=True,
            title=f"{dim} 分布统计"
        )
        
        # 将图表放入对应的列中
        with chart_cols[i]:
            st.plotly_chart(fig, use_container_width=True)

# ==========================================
# 6. 原始数据展开查看
# ==========================================
with st.expander("点击查看原始数据表格"):
    st.dataframe(df)
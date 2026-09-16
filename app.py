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
# 2. 动态数据通道
# ==========================================
st.sidebar.header("📁 数据通道管理")
uploaded_file = st.sidebar.file_uploader(
    "上传新的数据文件（支持 CSV 或 Excel）", 
    type=["csv", "xlsx", "xls"]
)

def load_default_data():
    try:
        xls = pd.ExcelFile("data.xlsx")
        all_dfs = []
        for sheet in xls.sheet_names:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df['来源工作表'] = sheet
            all_dfs.append(temp_df)
        return pd.concat(all_dfs, ignore_index=True)
    except Exception as e:
        raise RuntimeError(f"本地文件 data.xlsx 读取失败，请检查文件是否存在。原始错误：{e}")

def get_data():
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        if uploaded_file.name.endswith('.csv'):
            try:
                df_csv = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
            except UnicodeDecodeError:
                df_csv = pd.read_csv(io.BytesIO(file_bytes), encoding='gbk')
            df_csv['来源工作表'] = '上传的CSV文件'
            return df_csv
        else:
            try:
                xls = pd.ExcelFile(io.BytesIO(file_bytes))
                all_dfs = []
                for sheet in xls.sheet_names:
                    temp_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
                    temp_df['来源工作表'] = sheet
                    all_dfs.append(temp_df)
                return pd.concat(all_dfs, ignore_index=True)
            except Exception as e:
                st.sidebar.error(f"⚠️ 读取Excel失败！请确保文件未被其他软件打开。错误：{e}")
                st.stop()
    else:
        return load_default_data()

# ==========================================
# 3. 数据加载
# ==========================================
df = pd.DataFrame()

try:
    df = get_data()
    st.sidebar.success(f"数据加载成功！共合并了 {len(df)} 条记录")
except Exception as e:
    st.error(f"读取文件失败，请检查文件名和路径。报错信息：{e}")
    st.stop()

# ==========================================
# 4. 侧边栏数据源筛选
# ==========================================
st.sidebar.header("⚙️ 数据源筛选")
filtered_df = df.copy()

if '来源工作表' in filtered_df.columns:
    sheet_options = list(filtered_df['来源工作表'].dropna().unique())
    selected_sheets = st.sidebar.multiselect(
        "选择来源表格（可多选）：",
        options=sheet_options,
        default=sheet_options
    )
    if len(selected_sheets) == 0:
        st.sidebar.warning("⚠️ 请至少选择一个来源表格！")
        st.stop()
    filtered_df = filtered_df[filtered_df['来源工作表'].isin(selected_sheets)]

# ==========================================
# 5. 核心配置区
# ==========================================
CATEGORY_COL = '来源工作表' 

COMMON_COLS = {
    '片区': '片区',
    '蝶城': '蝶城'
}

TAB_CONFIG = {
    '防汛': {
        'risk': '防汛风险等级',       
        'point': '项目名称',          
        'desc': '风险描述'            
    },
    '消防': {
        'risk': '消防风险等级',
        'point': '项目名称',
        'desc': '风险描述'
    },
    '治安': {
        'risk': '治安风险等级',
        'point': '项目名称',
        'desc': '风险描述'
    }
}

# ==========================================
# 6. 核心渲染函数
# ==========================================
def render_category_dashboard(category_name, data_df, config):
    if CATEGORY_COL not in data_df.columns:
        st.error(f"⚠️ 表格中未找到分类列 '{CATEGORY_COL}'，无法精准过滤 {category_name} 数据。")
        return

    category_df = data_df[data_df[CATEGORY_COL] == category_name].copy()
        
    if len(category_df) == 0:
        st.info(f"💡 当前侧边栏筛选条件下暂无 {category_name} 相关数据（可能未勾选对应表格）。")
        return

    risk_col = config['risk']
    point_col = config['point']
    desc_col = config['desc']

    # 指标卡片
    col1, col2, col3 = st.columns(3)
    col1.metric(f"{category_name}项目总数", len(category_df))
    col2.metric("数据来源表格数", 1)
    
    if risk_col in category_df.columns:
        high_risk_mask = category_df[risk_col].astype(str).str.contains('高风险', na=False)
        high_risk_count = len(category_df[high_risk_mask])
        col3.metric("🚨 高风险项目数", high_risk_count, delta="需重点关注", delta_color="inverse")
    else:
        col3.metric("🚨 高风险项目数", "未知", delta=f"缺少 {risk_col} 列", delta_color="off")
    
    st.markdown("---")

    # ================= 修改点：取消片区分布统计，仅保留蝶城 =================
    target_dims = [col for col in [COMMON_COLS['蝶城']] if col in category_df.columns]
    # 如果之后“蝶城”也不想看了，把上面那行换成 target_dims = [] 即可
    if target_dims:
        chart_cols = st.columns(len(target_dims))
        for i, dim in enumerate(target_dims):
            chart_data = category_df[dim].value_counts().reset_index()
            chart_data.columns = ['分类', '数量']
            fig = px.bar(
                chart_data, x='分类', y='数量', color='分类', text_auto=True,
                title=f"{category_name} - {dim} 分布统计"
            )
            with chart_cols[i]:
                st.plotly_chart(fig, use_container_width=True)
    # =====================================================================

    # 高风险点位详细清单
    st.markdown("---")
    st.subheader(f"🚨 {category_name} - 高风险点位详细清单")

    if point_col not in category_df.columns or desc_col not in category_df.columns or risk_col not in category_df.columns:
        st.info(f"💡 提示：缺少 {point_col}、{desc_col} 或 {risk_col} 列，无法生成详细清单。请修改代码中的 TAB_CONFIG。")
    else:
        display_cols = []
        for col in [COMMON_COLS['片区'], COMMON_COLS['蝶城'], risk_col, point_col, desc_col]:
            if col in category_df.columns:
                display_cols.append(col)
        
        detail_df = category_df[display_cols].copy()
        
        detail_df[desc_col] = detail_df[desc_col].fillna('暂无详细描述').astype(str)
        detail_df[point_col] = detail_df[point_col].fillna('未知点位').astype(str)

        detail_df = detail_df[detail_df[risk_col].astype(str).str.contains('高风险', na=False)]

        if len(detail_df) == 0:
            st.success(f"🎉 {category_name} 当前筛选条件下暂无高风险点位！")
        else:
            st.error(f"⚠️ {category_name} 当前筛选条件下存在 {len(detail_df)} 个高风险点位，请重点关注！")
            
            st.dataframe(
                detail_df, use_container_width=True,
                column_config={
                    desc_col: st.column_config.TextColumn("详细描述", width="large"),
                    point_col: st.column_config.TextColumn("点位名称", width="medium")
                }
            )

            csv = detail_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 下载 {category_name} 高风险清单 (CSV)",
                data=csv, file_name=f'{category_name}_高风险点位清单.csv',
                mime='text/csv', key=f'download_{category_name}'
            )

# ==========================================
# 7. 主界面：垂直排列三个独立板块
# ==========================================
with st.expander("🌊 防汛板块", expanded=True):
    render_category_dashboard("防汛", filtered_df, TAB_CONFIG['防汛'])

st.markdown("---")

with st.expander("🔥 消防板块", expanded=True):
    render_category_dashboard("消防", filtered_df, TAB_CONFIG['消防'])

st.markdown("---")

with st.expander("🚓 治安板块", expanded=True):
    render_category_dashboard("治安", filtered_df, TAB_CONFIG['治安'])

# ==========================================
# 8. 原始数据展开查看
# ==========================================
st.markdown("---")
with st.expander("点击查看当前筛选后的原始数据表格"):
    st.dataframe(filtered_df, use_container_width=True)
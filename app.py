import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================================
# 1. 页面基础设置
# ==========================================
st.set_page_config(page_title="项目风险看板", layout="wide")
st.title("🚨 成都城市营业部项目风险看板")

# ==========================================
# 2. 数据加载（仅本地 data.xlsx）
# ==========================================
def load_default_data():
    try:
        xls = pd.ExcelFile("data.xlsx")
        all_dfs = []
        for sheet in xls.sheet_names:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df['来源工作表'] = sheet
            all_dfs.append(temp_df)
        df = pd.concat(all_dfs, ignore_index=True)
        # 去除所有列名前后的空格，防止匹配失败
        df.columns = df.columns.str.strip() 
        return df
    except Exception as e:
        raise RuntimeError(f"本地文件 data.xlsx 读取失败，请检查文件是否存在。原始错误：{e}")

try:
    filtered_df = load_default_data()
except Exception as e:
    st.error(f"读取文件失败，请检查文件名和路径。报错信息：{e}")
    st.stop()

# ==========================================
# 3. 核心配置区
# ==========================================
CATEGORY_COL = '来源工作表'

COMMON_COLS = {
    '片区': '片区',
    '蝶城': '蝶城' 
}

STREET_KEYWORD = '街区住宅' 

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
# 4. 核心渲染函数
# ==========================================
def render_category_dashboard(category_name, data_df, config):
    if CATEGORY_COL not in data_df.columns:
        st.error(f"⚠️ 表格中未找到分类列 '{CATEGORY_COL}'，无法精准过滤 {category_name} 数据。")
        return

    category_df = data_df[data_df[CATEGORY_COL] == category_name].copy()

    if len(category_df) == 0:
        st.info(f"💡 当前数据中暂无 {category_name} 相关数据。")
        return

    risk_col = config['risk']
    point_col = config['point']
    desc_col = config['desc']
    dim_col = COMMON_COLS['蝶城']

    # ================= 核心过滤：只保留高风险数据 =================
    if risk_col not in category_df.columns:
        st.warning(f"⚠️ 当前数据缺少 '{risk_col}' 列，无法进行高风险过滤。")
        return

    high_risk_df = category_df[category_df[risk_col].astype(str).str.contains('高风险', na=False)].copy()
    
    if len(high_risk_df) == 0:
        st.success(f"🎉 {category_name} 当前暂无高风险数据！")
        return

    # 在高风险数据中，分离“蝶城”与“街区住宅”用于指标卡
    if dim_col in high_risk_df.columns:
        is_street = high_risk_df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)
        high_risk_city = high_risk_df[~is_street]    # 高风险蝶城数据
        high_risk_street = high_risk_df[is_street]   # 高风险街区住宅数据
    else:
        high_risk_city = high_risk_df
        high_risk_street = pd.DataFrame()

    # ================= 1. 指标卡片（仅统计高风险） =================
    col1, col2, col3 = st.columns(3)
    col1.metric("🚨 高风险项目总数", len(high_risk_df))
    
    if dim_col in high_risk_city.columns:
        # 统计高风险蝶城项目数量
        col2.metric("🚨 高风险蝶城项目数", len(high_risk_city)) 
    else:
        col2.metric("🚨 高风险蝶城项目数", "未知")
        
    col3.metric("🚨 高风险街区住宅项目数", len(high_risk_street))

    st.markdown("---")

    # ================= 2. 高风险项目分布图表（含蝶城与街区住宅） =================
    if dim_col in high_risk_df.columns and len(high_risk_df) > 0:
        chart_df = high_risk_df.copy()
        
        # 处理空值：如果“蝶城”列是空值，填入“街区住宅项目”，使其能作为独立分类展示
        chart_df[dim_col] = chart_df[dim_col].fillna('街区住宅项目')
        
        # 统计数量
        chart_data = chart_df[dim_col].value_counts().reset_index()
        chart_data.columns = ['分类', '数量']
        
        fig = px.bar(
            chart_data, x='分类', y='数量', color='分类', text_auto=True,
            title=f"{category_name} - 高风险项目分布统计（含街区住宅）"
        )
        st.plotly_chart(fig, use_container_width=True)

    # ================= 3. 高风险点位详细清单 =================
    st.markdown("---")
    st.subheader(f"🚨 {category_name} - 高风险点位详细清单")

    if point_col not in high_risk_df.columns or desc_col not in high_risk_df.columns:
        st.info(f"💡 提示：缺少 {point_col} 或 {desc_col} 列，无法生成详细清单。")
    else:
        display_cols = []
        for col in [COMMON_COLS['片区'], COMMON_COLS['蝶城'], risk_col, point_col, desc_col]:
            if col in high_risk_df.columns:
                display_cols.append(col)

        # 切换按钮，用于明细清单的筛选
        view_type = st.radio(
            "选择查看范围：", 
            ["全部高风险项目", "仅看高风险蝶城项目", "仅看高风险街区住宅项目"], 
            horizontal=True, key=f"radio_{category_name}"
        )
        
        detail_df = high_risk_df[display_cols].copy()
        
        if dim_col in detail_df.columns:
            if view_type == "仅看高风险蝶城项目":
                detail_df = detail_df[~detail_df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)]
            elif view_type == "仅看高风险街区住宅项目":
                detail_df = detail_df[detail_df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)]

        detail_df[desc_col] = detail_df[desc_col].fillna('暂无详细描述').astype(str)
        detail_df[point_col] = detail_df[point_col].fillna('未知点位').astype(str)

        if len(detail_df) == 0:
            st.info(f"💡 {view_type} 条件下暂无数据。")
        else:
            st.error(f"⚠️ 当前筛选条件下存在 {len(detail_df)} 个高风险点位，请重点关注！")
            
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
# 5. 主界面：垂直排列三个独立板块
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
# 6. 原始数据展开查看
# ==========================================
st.markdown("---")
with st.expander("点击查看原始数据表格"):
    st.dataframe(filtered_df, use_container_width=True)
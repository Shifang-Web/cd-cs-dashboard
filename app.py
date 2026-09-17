import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import io
import os
import datetime
import numpy as np

# ==========================================
# 1. 页面基础设置
# ==========================================
st.set_page_config(page_title="项目风险看板", layout="wide")
st.title("🚨 成都城市营业部项目风险看板")

try:
    mtime = os.path.getmtime("data.xlsx")
    st.caption(f"📅 数据文件最后更新：{datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}")
except Exception:
    pass

# ==========================================
# 2. 配置读取
# ==========================================
def get_secret(key, default):
    try:
        return st.secrets["qweather"][key]
    except Exception:
        return default

QWEATHER_KEY = get_secret("key", "dc8db5bb8e1049d6addb14bef20e2bb5")
QWEATHER_API_HOST = get_secret("host", "https://nd5khxx2t4.re.qweatherapi.com/weatheralert/v1/current")
QWEATHER_BASE = QWEATHER_API_HOST.rsplit('/weatheralert', 1)[0]

# ==========================================
# 3. 数据加载
# ==========================================
@st.cache_data(ttl=600)
def load_default_data():
    try:
        xls = pd.ExcelFile("data.xlsx")
        all_dfs = []
        for sheet in xls.sheet_names:
            temp_df = pd.read_excel(xls, sheet_name=sheet)
            temp_df['来源工作表'] = sheet
            all_dfs.append(temp_df)
        df = pd.concat(all_dfs, ignore_index=True)
        df.columns = df.columns.str.strip().str.replace('\u3000', '')
        return df
    except Exception as e:
        raise RuntimeError(f"本地文件 data.xlsx 读取失败，请检查文件是否存在。原始错误：{e}")

try:
    filtered_df = load_default_data()
except Exception as e:
    st.error(f"读取文件失败，请检查文件名和路径。报错信息：{e}")
    st.stop()

# ==========================================
# 4. 核心配置区
# ==========================================
CATEGORY_COL = '来源工作表'

COMMON_COLS = {
    '片区': '片区',
    '蝶城': '蝶城',
    '经度': '经度',
    '纬度': '纬度'
}

STREET_KEYWORD = '街区住宅'

TAB_CONFIG = {
    '防汛': {'risk': '防汛风险等级', 'point': '项目名称', 'desc': '风险描述'},
    '消防': {'risk': '消防风险等级', 'point': '项目名称', 'desc': '风险描述'},
    '治安': {'risk': '治安风险等级', 'point': '项目名称', 'desc': '风险描述'}
}

SEVERITY_MAP = {
    'Extreme': ('红色', '#D0021B'),
    'Severe':  ('橙色', '#F5A623'),
    'Moderate':('黄色', '#F5A623'),
    'Minor':   ('蓝色', '#4A90E2'),
    'Unknown': ('未知', '#9B9B9B'),
    '红色':    ('红色', '#D0021B'),
    '橙色':    ('橙色', '#F5A623'),
    '黄色':    ('黄色', '#F5A623'),
    '蓝色':    ('蓝色', '#4A90E2'),
}

CITY_MAP_CONFIG = {
    '成都': {"center": {"lat": 30.67, "lon": 104.06}, "zoom": 9},
    '昆明': {"center": {"lat": 25.04, "lon": 102.71}, "zoom": 9},
    '西昌': {"center": {"lat": 27.89, "lon": 102.26}, "zoom": 10},
}

def weather_emoji(text):
    if not text: return '🌡️'
    if '雷' in text: return '⛈️'
    if '暴' in text: return '🌧️'
    if '大雨' in text or '中雨' in text: return '🌧️'
    if '雨' in text: return '🌦️'
    if '雪' in text: return '❄️'
    if '雾' in text or '霾' in text: return '🌫️'
    if '阴' in text: return '☁️'
    if '多云' in text: return '⛅'
    if '晴' in text: return '☀️'
    return '🌡️'

# ==========================================
# 5. 工具函数
# ==========================================
def extract_city(row):
    dim_col = COMMON_COLS['蝶城']
    dit = str(row.get(dim_col, '') or '')
    name = str(row.get('项目名称', '') or '')

    if dit.startswith("CD") or "成都" in dit:
        return "成都"
    if dit.startswith("KM") or "昆明" in dit:
        return "昆明"
    if "西昌" in dit:
        return "西昌"

    if "西昌" in name or "邛海" in name:
        return "西昌"
    if "昆明" in name:
        return "昆明"
    if "成都" in name or "蓉" in name or "锦江" in name or "青羊" in name or "成华" in name or "武侯" in name or "金牛" in name or "龙泉驿" in name or "郫都" in name or "天府新区" in name:
        return "成都"

    return "其他"

@st.cache_data(ttl=1800)
def get_weather_warnings_cached(lat, lon):
    url = f"{QWEATHER_API_HOST}/{lat}/{lon}"
    params = {'key': QWEATHER_KEY, 'lang': 'zh'}
    try:
        resp = requests.get(url, params=params, timeout=10)
        data = resp.json()
        if 'error' in data:
            err = data['error']
            return None, f"API错误：{err.get('title', '未知')} - {err.get('detail', '')}"
        if 'alerts' in data:
            return data['alerts'], None
        if 'code' in data:
            if data['code'] == '200':
                return data.get('warning', []), None
            return None, f"API返回码：{data['code']}"
        return None, f"未知响应：{data}"
    except Exception as e:
        return None, f"请求失败：{e}"

@st.cache_data(ttl=600)
def get_current_weather_cached(lat, lon):
    candidates = [
        (f"{QWEATHER_BASE}/weather/v1/now/{lat}/{lon}", {}),
        (f"{QWEATHER_BASE}/v7/weather/now", {'location': f'{lon},{lat}'}),
    ]
    for url, extra_params in candidates:
        params = {'key': QWEATHER_KEY, 'lang': 'zh'}
        params.update(extra_params)
        try:
            resp = requests.get(url, params=params, timeout=10)
            data = resp.json()
            if 'error' in data:
                continue
            code = data.get('code')
            if code not in (None, '200'):
                continue
            now = data.get('now') or data.get('current') or data.get('weather')
            if not now or 'temp' not in now:
                continue
            return now, None
        except Exception:
            continue
    return None, "无法获取天气数据（端点可能已变更）"

# ==========================================
# 6. 板块数据处理
# ==========================================
def get_category_high_risk_df(category_name, data_df, config):
    if CATEGORY_COL not in data_df.columns:
        return pd.DataFrame(), 0, 0
    category_df = data_df[data_df[CATEGORY_COL] == category_name].copy()
    if len(category_df) == 0:
        return pd.DataFrame(), 0, 0
    risk_col = config['risk']
    dim_col = COMMON_COLS['蝶城']
    if risk_col not in category_df.columns:
        return pd.DataFrame(), 0, 0
    high_risk_df = category_df[category_df[risk_col].astype(str).str.contains('高风险', na=False)].copy()
    if len(high_risk_df) == 0:
        return pd.DataFrame(), 0, 0
    if dim_col in high_risk_df.columns:
        is_street = high_risk_df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)
        n_city = len(high_risk_df[~is_street])
        n_street = len(high_risk_df[is_street])
    else:
        n_city = len(high_risk_df)
        n_street = 0
    return high_risk_df, n_city, n_street

hr_fangxun, n_city_fx, n_street_fx = get_category_high_risk_df("防汛", filtered_df, TAB_CONFIG['防汛'])
hr_xiaofang, n_city_xf, n_street_xf = get_category_high_risk_df("消防", filtered_df, TAB_CONFIG['消防'])
hr_zhian, n_city_za, n_street_za = get_category_high_risk_df("治安", filtered_df, TAB_CONFIG['治安'])

# ==========================================
# 7. 城市代表坐标
# ==========================================
def build_city_representatives():
    all_hr = pd.concat([hr_fangxun, hr_xiaofang, hr_zhian], ignore_index=True)
    dim_col = COMMON_COLS['蝶城']
    lon_col = COMMON_COLS['经度']
    lat_col = COMMON_COLS['纬度']

    if len(all_hr) == 0:
        return {}, "三个板块均无高风险数据"

    missing_cols = []
    if dim_col not in all_hr.columns:
        missing_cols.append(f"蝶城列（期望列名'{dim_col}'）")
    if lon_col not in all_hr.columns:
        missing_cols.append(f"经度列（期望列名'{lon_col}'）")
    if lat_col not in all_hr.columns:
        missing_cols.append(f"纬度列（期望列名'{lat_col}'）")

    if missing_cols:
        return {}, (
            f"❌ 缺少以下列：{'、'.join(missing_cols)}\n\n"
            f"📋 当前数据实际列名：{list(all_hr.columns)}"
        )

    tmp = all_hr.copy()
    tmp['_city'] = tmp.apply(extract_city, axis=1)
    tmp[lon_col] = pd.to_numeric(tmp[lon_col], errors='coerce')
    tmp[lat_col] = pd.to_numeric(tmp[lat_col], errors='coerce')
    tmp = tmp.dropna(subset=[lon_col, lat_col])
    if len(tmp) == 0:
        return {}, "❌ 经纬度数据全部为空或无法转为数字"

    city_reps = {}
    for city, group in tmp.groupby('_city'):
        if city == "其他":
            continue
        first = group.iloc[0]
        city_reps[city] = (first[lat_col], first[lon_col])
    if not city_reps:
        return {}, "❌ 所有项目的城市都无法识别"
    return city_reps, None

# ==========================================
# 8. 天气预警板块
# ==========================================
def render_global_weather_section():
    st.subheader("🌦️ 天气预警（全区域）")
    city_reps, err = build_city_representatives()
    if err:
        st.warning(err)
        return set()

    alert_cities = set()
    cols = st.columns(min(len(city_reps), 3))
    for i, (city, (lat, lon)) in enumerate(sorted(city_reps.items())):
        warnings, error = get_weather_warnings_cached(lat, lon)
        with cols[i % 3]:
            if error is not None:
                st.warning(f"⚠️ {city} 天气预警获取失败：{error}")
                continue
            if len(warnings) == 0:
                st.success(f"✅ {city} 当前无天气预警")
                continue
            alert_cities.add(city)
            for warn in warnings[:3]:
                severity = warn.get('severity', 'Unknown')
                title = warn.get('title', '天气预警')
                text = warn.get('description', '')
                event = warn.get('event', '')
                sent = warn.get('sent', '')
                instruction = warn.get('instruction', '')
                level_cn, color = SEVERITY_MAP.get(severity, ('未知', '#9B9B9B'))
                instruction_html = ""
                if instruction:
                    instruction_html = f'''
                        <details style="margin-top: 4px;">
                            <summary style="font-size: 12px; color: #555; cursor: pointer;">展开防御指南</summary>
                            <div style="font-size: 12px; color: #555; margin-top: 4px;">{instruction}</div>
                        </details>
                    '''
                st.markdown(f"""
                    <div style="border-left: 5px solid {color}; background: #F7F9FC; padding: 10px 14px; border-radius: 6px; margin-bottom: 8px;">
                        <div style="font-weight: 600; font-size: 14px; color: {color};">{city} · {title}</div>
                        <div style="font-size: 12px; color: #666; margin: 4px 0;">等级：{level_cn} · 类型：{event}</div>
                        <div style="font-size: 12px; color: #888; margin-top: 6px;">发布时间：{sent}</div>
                        <details style="margin-top: 6px;">
                            <summary style="font-size: 12px; color: #555; cursor: pointer;">展开详细描述</summary>
                            <div style="font-size: 12px; color: #555; margin-top: 4px;">{text}</div>
                        </details>
                        {instruction_html}
                    </div>
                """, unsafe_allow_html=True)

    if alert_cities:
        st.warning(f"⚠️ 当前有天气预警的城市：{', '.join(sorted(alert_cities))}")
    else:
        st.success("✅ 所有涉及城市当前均无天气预警")
    return alert_cities

# ==========================================
# 9. 当日城市天气板块
# ==========================================
def render_global_weather_now_section():
    st.subheader("🌤️ 当日城市天气")
    city_reps, err = build_city_representatives()
    if err:
        st.info(f"💡 {err}")
        return

    cols = st.columns(min(len(city_reps), 3))
    for i, (city, (lat, lon)) in enumerate(sorted(city_reps.items())):
        weather, error = get_current_weather_cached(lat, lon)
        with cols[i % 3]:
            if error is not None or weather is None:
                st.warning(f"⚠️ {city} 天气获取失败：{error or '未知'}")
                continue
            temp = weather.get('temp', '--')
            feels = weather.get('feelsLike', '--')
            text = weather.get('text', '--')
            humidity = weather.get('humidity', '--')
            wind_dir = weather.get('windDir', '--')
            wind_scale = weather.get('windScale', '--')
            obs_time = weather.get('obsTime', '')
            icon = weather_emoji(text)
            time_str = obs_time[11:16] if len(obs_time) >= 16 else obs_time
            st.markdown(f"""
                <div style="background: linear-gradient(135deg, #E8F4FD 0%, #F7F9FC 100%); border: 1px solid #D6E4F0; border-radius: 10px; padding: 16px 18px; margin-bottom: 10px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div style="font-size: 16px; font-weight: 600; color: #2C5282;">🏙️ {city}</div>
                        <div style="font-size: 32px;">{icon}</div>
                    </div>
                    <div style="font-size: 28px; font-weight: 700; color: #1A365D; margin: 6px 0;">
                        {temp}<span style="font-size: 16px; font-weight: 400;">°C</span>
                        <span style="font-size: 14px; font-weight: 400; color: #4A5568; margin-left: 8px;">{text}</span>
                    </div>
                    <div style="font-size: 12px; color: #4A5568; margin-top: 6px; line-height: 1.8;">
                        🌡️ 体感温度：{feels}°C<br/>
                        💧 相对湿度：{humidity}%<br/>
                        🌬️ 风向风力：{wind_dir} {wind_scale}级
                    </div>
                    <div style="font-size: 11px; color: #A0AEC0; margin-top: 8px; text-align: right;">
                        更新时间：{time_str}
                    </div>
                </div>
            """, unsafe_allow_html=True)

# ==========================================
# 10. 渲染两个独立板块
# ==========================================
with st.expander("🌦️ 天气预警板块", expanded=True):
    alert_cities_global = render_global_weather_section()

st.markdown("---")

with st.expander("🌤️ 当日城市天气板块", expanded=True):
    render_global_weather_now_section()

st.markdown("---")

# ==========================================
# 11. 区域绘制函数（支持"保和红/其他黄"和"全红"两种模式）
# ==========================================
def add_risk_zones_by_dit(fig, map_df, lon_col, lat_col, dim_col, highlight_baohe=True):
    """
    按蝶城分组画区域圈。
    highlight_baohe=True：保和红圈 + 其他黄圈（防汛板块）
    highlight_baohe=False：全部红圈（消防、治安板块）
    """
    if dim_col not in map_df.columns:
        return 0, ["地图数据中缺少蝶城列"]

    df_dit = map_df[~map_df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)].copy()
    if len(df_dit) == 0:
        return 0, ["没有可用的蝶城项目"]

    zones_drawn = 0
    errors = []

    if hasattr(go, 'Scattermap'):
        scatter_cls = go.Scattermap
    elif hasattr(go, 'Scattermapbox'):
        scatter_cls = go.Scattermapbox
    else:
        return 0, ["当前 Plotly 版本不支持 Scattermap"]

    COLOR_HIGHLIGHT = {'line': '#FF0000', 'fill': 'rgba(255, 0, 0, 0.20)', 'text': '#B00020'}
    COLOR_NORMAL = {'line': '#FFB800', 'fill': 'rgba(255, 184, 0, 0.22)', 'text': '#8B6508'}
    COLOR_ALL_RED = {'line': '#FF0000', 'fill': 'rgba(255, 0, 0, 0.18)', 'text': '#B00020'}

    for dit_name, group in df_dit.groupby(dim_col):
        coords = group[[lon_col, lat_col]].dropna().values
        if len(coords) < 1:
            continue

        center_lon = coords[:, 0].mean()
        center_lat = coords[:, 1].mean()

        if len(coords) >= 2:
            dists = np.sqrt((coords[:, 0] - center_lon)**2 + (coords[:, 1] - center_lat)**2)
            radius = max(dists.max() * 1.6, 0.01)
        else:
            radius = 0.015

        theta = np.linspace(0, 2 * np.pi, 100)
        circle_lon = center_lon + radius * np.cos(theta)
        circle_lat = center_lat + radius * np.sin(theta)

        # 配色选择
        if not highlight_baohe:
            color_cfg = COLOR_ALL_RED
        else:
            color_cfg = COLOR_HIGHLIGHT if '保和' in str(dit_name) else COLOR_NORMAL

        try:
            fig.add_trace(scatter_cls(
                lon=circle_lon,
                lat=circle_lat,
                mode='lines',
                line=dict(color=color_cfg['line'], width=3),
                fill='toself',
                fillcolor=color_cfg['fill'],
                hoverinfo='skip',
                showlegend=False,
                name=str(dit_name)
            ))

            fig.add_trace(scatter_cls(
                lon=[center_lon],
                lat=[center_lat],
                mode='text',
                text=[str(dit_name)],
                textfont=dict(size=13, color=color_cfg['text'],
                              family='Microsoft YaHei, SimHei, sans-serif'),
                textposition='middle center',
                hoverinfo='skip',
                showlegend=False,
                name=''
            ))

            zones_drawn += 1
        except Exception as e:
            errors.append(f"{dit_name}: {e}")

    return zones_drawn, errors

# ==========================================
# 12. 单城市地图渲染函数
# ==========================================
def render_city_map(city_name, city_df, category_name, config, alert_cities, highlight_baohe=True):
    risk_col = config['risk']
    point_col = config['point']
    dim_col = COMMON_COLS['蝶城']
    lon_col = COMMON_COLS['经度']
    lat_col = COMMON_COLS['纬度']

    if len(city_df) == 0:
        st.info(f"💡 {city_name} 无高风险项目")
        return

    if lon_col not in city_df.columns or lat_col not in city_df.columns:
        st.info(f"💡 {city_name} 数据缺少经纬度列")
        return

    map_df = city_df.copy()
    map_df[dim_col] = map_df[dim_col].fillna('街区住宅项目')
    map_df[lon_col] = pd.to_numeric(map_df[lon_col], errors='coerce')
    map_df[lat_col] = pd.to_numeric(map_df[lat_col], errors='coerce')
    map_df = map_df.dropna(subset=[lon_col, lat_col])

    if len(map_df) == 0:
        st.warning(f"⚠️ {city_name} 无有效坐标数据")
        return

    if map_df.iloc[0][lon_col] < 50:
        map_df[[lon_col, lat_col]] = map_df[[lat_col, lon_col]]

    cfg = CITY_MAP_CONFIG.get(city_name, {"center": {"lat": 30.67, "lon": 104.06}, "zoom": 9})

    if hasattr(go, 'Scattermap'):
        scatter_cls = go.Scattermap
        map_layout_key = 'map'
    elif hasattr(go, 'Scattermapbox'):
        scatter_cls = go.Scattermapbox
        map_layout_key = 'mapbox'
    else:
        st.error("当前 Plotly 版本不支持 Scattermap")
        return

    def make_hover(df):
        return [
            f"<b>{row[point_col]}</b><br>蝶城：{row[dim_col]}<br>{risk_col}：{row[risk_col]}"
            for _, row in df.iterrows()
        ]

    with st.spinner(f"正在加载 {city_name} 地图..."):
        fig = go.Figure()

        if highlight_baohe:
            # 【防汛板块】保和红点大、其他黄点小
            is_baohe = map_df[dim_col].astype(str).str.contains('保和', na=False)
            baohe_df = map_df[is_baohe].copy()
            other_df = map_df[~is_baohe].copy()

            if len(other_df) > 0:
                fig.add_trace(scatter_cls(
                    lon=other_df[lon_col],
                    lat=other_df[lat_col],
                    mode='markers+text',
                    marker=dict(size=9, color='#FFC107', opacity=0.9),
                    text=other_df[point_col].astype(str).tolist(),
                    textposition='top center',
                    textfont=dict(size=8, color='#666'),
                    hovertext=make_hover(other_df),
                    hoverinfo='text',
                    showlegend=False,
                    name='其他蝶城'
                ))

            if len(baohe_df) > 0:
                fig.add_trace(scatter_cls(
                    lon=baohe_df[lon_col],
                    lat=baohe_df[lat_col],
                    mode='markers+text',
                    marker=dict(size=18, color='#FF0000', opacity=0.9),
                    text=baohe_df[point_col].astype(str).tolist(),
                    textposition='top center',
                    textfont=dict(size=10, color='#333'),
                    hovertext=make_hover(baohe_df),
                    hoverinfo='text',
                    showlegend=False,
                    name='保和蝶城'
                ))
        else:
            # 【消防/治安板块】全部红点
            if len(map_df) > 0:
                fig.add_trace(scatter_cls(
                    lon=map_df[lon_col],
                    lat=map_df[lat_col],
                    mode='markers+text',
                    marker=dict(size=18, color='#FF0000', opacity=0.9),
                    text=map_df[point_col].astype(str).tolist(),
                    textposition='top center',
                    textfont=dict(size=10, color='#333'),
                    hovertext=make_hover(map_df),
                    hoverinfo='text',
                    showlegend=False,
                    name='高风险项目'
                ))

        # 区域圈（跟随 highlight_baohe 开关）
        zones_drawn, errors = add_risk_zones_by_dit(
            fig, map_df, lon_col, lat_col, dim_col, highlight_baohe=highlight_baohe
        )
        if zones_drawn == 0 and errors:
            with st.expander(f"⚠️ {city_name} 高风险区域绘制提示"):
                for e in errors:
                    st.write(f"- {e}")

        # 预警图标
        if city_name in alert_cities:
            fig.add_trace(scatter_cls(
                lat=[map_df.iloc[0][lat_col]],
                lon=[map_df.iloc[0][lon_col]],
                mode='markers+text',
                marker=dict(size=10, color='rgba(255,215,0,0.001)'),
                text=['⚠️'],
                textfont=dict(size=40, color='#FFAA00'),
                textposition='top center',
                hoverinfo='skip',
                showlegend=False,
                name='预警'
            ))

        layout_kwargs = {
            "margin": {"r": 0, "t": 40, "l": 0, "b": 0},
            "height": 550,
            "title": f"{category_name} - {city_name} 高风险项目分布",
        }
        layout_kwargs[map_layout_key] = dict(
            style="carto-positron-nolabels",
            center=cfg["center"],
            zoom=cfg["zoom"]
        )
        fig.update_layout(**layout_kwargs)
        st.plotly_chart(fig, use_container_width=True)

# ==========================================
# 13. 板块渲染函数
# ==========================================
def render_category_dashboard(category_name, high_risk_df, n_city, n_street, config, alert_cities,
                              show_xichang=True, highlight_baohe=True):
    risk_col = config['risk']
    point_col = config['point']
    desc_col = config['desc']
    dim_col = COMMON_COLS['蝶城']

    if len(high_risk_df) == 0:
        st.success(f"🎉 {category_name} 当前暂无高风险数据！")
        return pd.DataFrame()

    col1, col2, col3 = st.columns(3)
    col1.metric("🚨 高风险项目总数", len(high_risk_df))
    col2.metric("🚨 高风险蝶城项目数", n_city)
    col3.metric("🚨 高风险街区住宅项目数", n_street)

    st.markdown("---")
    st.subheader(f"🗺️ {category_name} - 高风险项目地图分布")

    hr_copy = high_risk_df.copy()
    hr_copy['_city'] = hr_copy.apply(extract_city, axis=1)

    chengdu_df = hr_copy[hr_copy['_city'] == '成都']
    kunming_df = hr_copy[hr_copy['_city'] == '昆明']
    xichang_df = hr_copy[hr_copy['_city'] == '西昌']

    left_col, right_col = st.columns(2)
    with left_col:
        render_city_map("成都", chengdu_df, category_name, config, alert_cities,
                        highlight_baohe=highlight_baohe)
    with right_col:
        render_city_map("昆明", kunming_df, category_name, config, alert_cities,
                        highlight_baohe=highlight_baohe)

    if show_xichang and len(xichang_df) > 0:
        st.markdown(f"**🗺️ {category_name} - 西昌 高风险项目分布**")
        render_city_map("西昌", xichang_df, category_name, config, alert_cities,
                        highlight_baohe=highlight_baohe)

    st.markdown("---")
    st.subheader(f"🚨 {category_name} - 高风险点位详细清单")

    detail_df = pd.DataFrame()
    if point_col not in high_risk_df.columns or desc_col not in high_risk_df.columns:
        st.info(f"💡 提示：缺少 {point_col} 或 {desc_col} 列，无法生成详细清单。")
    else:
        display_cols = [col for col in [COMMON_COLS['片区'], COMMON_COLS['蝶城'], risk_col, point_col, desc_col]
                        if col in high_risk_df.columns]
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
            st.dataframe(detail_df, use_container_width=True, column_config={
                desc_col: st.column_config.TextColumn("详细描述", width="large"),
                point_col: st.column_config.TextColumn("点位名称", width="medium")
            })
            csv = detail_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 下载 {category_name} 高风险清单 (CSV)",
                data=csv, file_name=f'{category_name}_高风险点位清单.csv',
                mime='text/csv', key=f'download_{category_name}'
            )
    return detail_df

# ==========================================
# 14. 主界面：三个风险板块
# ==========================================
with st.expander("🌊 防汛板块", expanded=True):
    detail_fangxun = render_category_dashboard(
        "防汛", hr_fangxun, n_city_fx, n_street_fx, TAB_CONFIG['防汛'], alert_cities_global,
        show_xichang=True,
        highlight_baohe=True     # 防汛：保和红点/其他黄点
    )

st.markdown("---")

with st.expander("🔥 消防板块", expanded=True):
    detail_xiaofang = render_category_dashboard(
        "消防", hr_xiaofang, n_city_xf, n_street_xf, TAB_CONFIG['消防'], alert_cities_global,
        show_xichang=False,
        highlight_baohe=False    # 消防：全部红点
    )

st.markdown("---")

with st.expander("🚓 治安板块", expanded=True):
    detail_zhian = render_category_dashboard(
        "治安", hr_zhian, n_city_za, n_street_za, TAB_CONFIG['治安'], alert_cities_global,
        show_xichang=True,
        highlight_baohe=False    # 治安：全部红点
    )

# ==========================================
# 15. 一键导出完整报告
# ==========================================
st.markdown("---")
st.subheader("📤 导出完整风险报告")

def build_excel_report(sheets):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame({"提示": ["该板块暂无高风险数据"]}).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
    return buffer.getvalue()

excel_bytes = build_excel_report({
    "防汛高风险": detail_fangxun,
    "消防高风险": detail_xiaofang,
    "治安高风险": detail_zhian,
})

st.download_button(
    label="📥 一键导出三个板块高风险清单 (Excel 多 sheet)",
    data=excel_bytes,
    file_name=f'项目风险报告_{datetime.datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    key='download_full_report'
)

# ==========================================
# 16. 原始数据展开查看
# ==========================================
st.markdown("---")
with st.expander("点击查看原始数据表格"):
    st.dataframe(filtered_df, use_container_width=True)
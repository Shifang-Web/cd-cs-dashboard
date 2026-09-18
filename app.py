import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import io
import os
import datetime
import html as html_module
import numpy as np

try:
    from streamlit_autorefresh import st_autorefresh
    _HAS_AUTOREFRESH = True
except ImportError:
    _HAS_AUTOREFRESH = False

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

CATEGORY_EMOJI = {'防汛': '🌊', '消防': '🔥', '治安': '🚓'}

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
# 11. 板块数据准备函数（只算详情，不渲染）
# ==========================================
def build_category_detail(category_name, high_risk_df, config):
    """返回该板块的详情 DataFrame（用于横排指标和滚动清单）。"""
    risk_col = config['risk']
    point_col = config['point']
    desc_col = config['desc']

    if len(high_risk_df) == 0:
        return pd.DataFrame()

    if point_col not in high_risk_df.columns or desc_col not in high_risk_df.columns:
        return pd.DataFrame()

    display_cols = [col for col in [COMMON_COLS['片区'], COMMON_COLS['蝶城'], risk_col, point_col, desc_col]
                    if col in high_risk_df.columns]
    detail = high_risk_df[display_cols].copy()
    detail[desc_col] = detail[desc_col].fillna('暂无详细描述').astype(str)
    detail[point_col] = detail[point_col].fillna('未知点位').astype(str)
    return detail

# ==========================================
# 12. 微信联系人式清单：setInterval + scrollTop 驱动滚动
# ==========================================
def _esc(v):
    return html_module.escape(str(v) if v is not None else '')


def _build_contacts_doc(df, config, height=520, px_per_sec=22.0, paused=False):
    """微信联系人风格列表 + setInterval + scrollTop 逐帧驱动。"""
    risk_col = config['risk']
    point_col = config['point']
    desc_col = config['desc']
    dim_col = COMMON_COLS['蝶城']
    area_col = COMMON_COLS['片区']

    rows = []
    for _, row in df.iterrows():
        area = _esc(row.get(area_col, ''))
        dit = _esc(row.get(dim_col, ''))
        risk = _esc(row.get(risk_col, ''))
        point = _esc(row.get(point_col, ''))
        desc = _esc(row.get(desc_col, ''))
        avatar_char = point.strip()[:1] if point.strip() else '项'

        rows.append(
            '<div class="rk-row">'
            '  <div class="rk-avatar">' + avatar_char + '</div>'
            '  <div class="rk-main">'
            '    <div class="rk-line1">'
            f'      <span class="rk-name">{point}</span>'
            f'      <span class="rk-badge">{risk}</span>'
            '    </div>'
            f'    <div class="rk-line2">🏢 {dit} · 📍 {area}</div>'
            f'    <div class="rk-line3">{desc}</div>'
            '  </div>'
            '</div>'
        )

    if not rows:
        return "<!DOCTYPE html><html><body></body></html>"

    single_html = "".join(rows)
    auto_init = "false" if paused else "true"

    doc = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'><style>"
        "*{box-sizing:border-box;}"
        "html,body{margin:0;padding:0;background:#FFFFFF;"
        "font-family:-apple-system,'Segoe UI','Microsoft YaHei',sans-serif;}"
        ".rk-viewport{"
        f"height:{int(height)}px;"
        "overflow-y:auto;overflow-x:hidden;"
        "scrollbar-width:none;-ms-overflow-style:none;"
        "border:1px solid #E5E7EB;border-radius:10px;background:#FFFFFF;"
        "}"
        ".rk-viewport::-webkit-scrollbar{display:none;width:0;height:0;}"
        ".rk-row{"
        "display:flex;align-items:flex-start;gap:10px;"
        "padding:10px 12px;border-bottom:1px solid #F3F4F6;"
        "}"
        ".rk-avatar{"
        "width:38px;height:38px;border-radius:8px;flex:0 0 38px;"
        "background:#D0021B;color:#FFFFFF;"
        "display:flex;align-items:center;justify-content:center;"
        "font-weight:600;font-size:15px;"
        "}"
        ".rk-main{flex:1;min-width:0;}"
        ".rk-line1{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}"
        ".rk-name{font-weight:600;color:#111827;font-size:13.5px;"
        "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:70%;}"
        ".rk-badge{"
        "background:#FEE2E2;color:#B91C1C;font-size:10px;font-weight:600;"
        "padding:1px 6px;border-radius:4px;white-space:nowrap;"
        "}"
        ".rk-line2{color:#6B7280;font-size:11.5px;margin-top:2px;"
        "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}"
        ".rk-line3{color:#4B5563;font-size:12px;line-height:1.5;margin-top:3px;"
        "display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;"
        "overflow:hidden;}"
        "</style></head><body>"
        '<div class="rk-viewport" id="vp">'
        f'<div id="single">{single_html}</div>'
        f'<div>{single_html}</div>'
        '</div>'
        "<script>"
        "(function(){"
        "var vp=document.getElementById('vp');"
        "var single=document.getElementById('single');"
        f"var SPEED={float(px_per_sec)};"
        f"var AUTO={auto_init};"
        "var paused=false;"
        "var h=0;"
        "var y=0;"
        "var acc=0;"
        "var INTERVAL=30;"
        "function measure(){"
        "  var hh=single.offsetHeight;"
        "  if(hh>0)h=hh;"
        "}"
        "measure();"
        "setTimeout(measure,100);"
        "setTimeout(measure,400);"
        "setTimeout(measure,1000);"
        "window.addEventListener('resize',measure);"
        "setInterval(function(){"
        "  if(h<=0){measure();return;}"
        "  if(!AUTO||paused)return;"
        "  acc+=SPEED*INTERVAL/1000;"
        "  var step=Math.floor(acc);"
        "  if(step<=0)return;"
        "  acc-=step;"
        "  y+=step;"
        "  while(y>=h){y-=h;}"
        "  vp.scrollTop=y;"
        "},INTERVAL);"
        "vp.addEventListener('mouseenter',function(){paused=true;});"
        "vp.addEventListener('mouseleave',function(){paused=false;});"
        "})();"
        "</script>"
        "</body></html>"
    )
    return doc


def render_scrolling_list(df, config, height=520, seconds_per_item=2.5, paused=False):
    """条目高度约 90px，秒/条 换算为 px/秒。"""
    px_per_sec = max(12.0, 90.0 / max(0.5, float(seconds_per_item)))
    doc = _build_contacts_doc(df, config, height=height,
                              px_per_sec=px_per_sec, paused=paused)
    components.html(doc, height=int(height) + 4, scrolling=False)


def _filter_by_view(df, view_type):
    dim_col = COMMON_COLS['蝶城']
    if dim_col not in df.columns:
        return df
    if view_type == "仅看高风险蝶城项目":
        return df[~df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)]
    if view_type == "仅看高风险街区住宅项目":
        return df[df[dim_col].astype(str).str.contains(STREET_KEYWORD, na=False)]
    return df


# ==========================================
# 13. 顶部：三个板块横排指标卡
# ==========================================
def render_metric_card(title_emoji, title_text, total, n_city, n_street, accent_color="#D0021B"):
    """单个板块的横排指标卡。"""
    st.markdown(f"""
        <div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;
                    background:linear-gradient(135deg,#FFFFFF 0%,#F9FAFB 100%);
                    height:100%;">
            <div style="font-size:15px;font-weight:600;color:#1A202C;margin-bottom:10px;">
                {title_emoji} {title_text}板块
            </div>
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:8px 0;border-top:1px solid #F3F4F6;">
                <span style="color:#6B7280;font-size:12.5px;">🚨 高风险项目总数</span>
                <span style="color:{accent_color};font-size:18px;font-weight:700;">{total}</span>
            </div>
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:8px 0;border-top:1px solid #F3F4F6;">
                <span style="color:#6B7280;font-size:12.5px;">🏢 高风险蝶城项目</span>
                <span style="color:{accent_color};font-size:18px;font-weight:700;">{n_city}</span>
            </div>
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:8px 0;border-top:1px solid #F3F4F6;">
                <span style="color:#6B7280;font-size:12.5px;">🏘️ 高风险街区住宅</span>
                <span style="color:{accent_color};font-size:18px;font-weight:700;">{n_street}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)


st.subheader("📊 三大板块高风险指标")
metric_cols = st.columns(3)
with metric_cols[0]:
    render_metric_card("🌊", "防汛", len(hr_fangxun), n_city_fx, n_street_fx)
with metric_cols[1]:
    render_metric_card("🔥", "消防", len(hr_xiaofang), n_city_xf, n_street_xf)
with metric_cols[2]:
    render_metric_card("🚓", "治安", len(hr_zhian), n_city_za, n_street_za)

st.markdown("---")

# ==========================================
# 14. 三板块横向滚动清单
# ==========================================
def render_all_scrolling_boards(sections, height=520, default_speed=2.5):
    st.subheader("🚨 高风险点位清单（微信联系人式滚动）")

    valid = [(n, d, c) for n, d, c in sections if d is not None and len(d) > 0]
    if not valid:
        st.success("🎉 当前所有板块均无高风险点位数据。")
        return

    ctrl1, ctrl2, ctrl3 = st.columns([3, 1, 1.2])
    with ctrl1:
        view_type = st.radio(
            "查看范围：",
            ["全部高风险项目", "仅看高风险蝶城项目", "仅看高风险街区住宅项目"],
            horizontal=True,
            key="carousel_view_type",
        )
    with ctrl2:
        paused = st.toggle("⏸ 全部暂停", value=False, key="pause_scroll")
    with ctrl3:
        speed = st.slider(
            "滚动速度（秒/条）", min_value=0.8, max_value=6.0,
            value=float(default_speed), step=0.2, key="scroll_speed",
            help="每条滑过的时间越短，滚动越快",
        )

    processed = []
    for n, d, c in valid:
        df = _filter_by_view(d.copy(), view_type).reset_index(drop=True)
        if len(df) > 0:
            processed.append((n, df, c))

    if not processed:
        st.info(f"💡 在「{view_type}」条件下暂无高风险点位。")
        return

    board_cols = st.columns(len(processed))
    for i, (name, df, config) in enumerate(processed):
        with board_cols[i]:
            emoji = CATEGORY_EMOJI.get(name, "📋")
            st.markdown(
                f"<div style='font-size:15px;font-weight:600;color:#1A202C;"
                f"padding:2px 0 4px 0;'>"
                f"{emoji} {name}板块 · "
                f"<span style='color:#D0021B;'>{len(df)}</span> 个高风险点位"
                f"</div>",
                unsafe_allow_html=True,
            )

            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label=f"📥 下载 {name} 清单 (CSV)",
                data=csv,
                file_name=f"{name}_高风险点位清单.csv",
                mime="text/csv",
                key=f"download_board_{name}",
                use_container_width=True,
            )

            render_scrolling_list(
                df, config,
                height=height,
                seconds_per_item=speed,
                paused=paused,
            )

            st.caption("💡 列表持续向上滚动，鼠标悬停可暂停")


# ==========================================
# 15. 生成三个板块详情数据 + 渲染滚动清单
# ==========================================
detail_fangxun = build_category_detail("防汛", hr_fangxun, TAB_CONFIG['防汛'])
detail_xiaofang = build_category_detail("消防", hr_xiaofang, TAB_CONFIG['消防'])
detail_zhian = build_category_detail("治安", hr_zhian, TAB_CONFIG['治安'])

render_all_scrolling_boards(
    [
        ("防汛", detail_fangxun, TAB_CONFIG['防汛']),
        ("消防", detail_xiaofang, TAB_CONFIG['消防']),
        ("治安", detail_zhian, TAB_CONFIG['治安']),
    ],
    height=520,
    default_speed=2.5,
)

# ==========================================
# 16. 一键导出完整报告
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
# 17. 原始数据展开查看
# ==========================================
st.markdown("---")
with st.expander("点击查看原始数据表格"):
    st.dataframe(filtered_df, use_container_width=True)
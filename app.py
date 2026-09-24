import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import io
import os
import time
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
st.title("🚨 成都城市营业部项目风险看板 🚨")

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


def get_wecom_webhook_default():
    """从 secrets 读取企微机器人 Webhook（可选）。"""
    try:
        return st.secrets["wecom"]["webhook"]
    except Exception:
        return ""


# ==========================================
# 3. 侧边栏：企微推送设置 + 天气自动刷新设置
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ 企业微信推送设置")
    WECOM_WEBHOOK = st.text_input(
        "群机器人 Webhook 地址",
        value=get_wecom_webhook_default(),
        type="password",
        help="企业微信群 → 群机器人 → 复制 Webhook 地址",
    )
    AUTO_PUSH = st.checkbox("启用定时自动推送", value=False)
    PUSH_INTERVAL_MIN = st.selectbox(
        "推送间隔", [15, 30, 60, 120, 180, 360, 720], index=2,
        format_func=lambda m: f"{m} 分钟",
        disabled=not AUTO_PUSH,
    )
    ONLY_WHEN_ALERT = st.checkbox("仅在有预警时推送", value=True, disabled=not AUTO_PUSH)
    BRIEF_PUSH = st.checkbox("推送精简版（不含详情/防御指南）", value=False)

    st.markdown("---")
    # ---------- 天气自动刷新 ----------
    st.markdown("### 🌦️ 天气自动刷新")

    WEATHER_AUTO_REFRESH = st.checkbox(
        "启用天气自动刷新", value=True, key="cfg_weather_auto_refresh",
        help="按设定间隔自动重新拉取天气预警与实时天气数据",
    )
    WEATHER_REFRESH_MIN = st.selectbox(
        "天气刷新间隔",
        [5, 10, 15, 20, 30, 60],
        index=2,  # 默认 15 分钟
        format_func=lambda m: f"{m} 分钟",
        disabled=not WEATHER_AUTO_REFRESH,
        key="cfg_weather_refresh_min",
        help="建议 15 分钟：兼顾预警时效性与和风天气免费版每日调用配额",
    )
    ALERT_SOUND = st.checkbox(
        "新预警提示音", value=True,
        disabled=not WEATHER_AUTO_REFRESH, key="cfg_alert_sound",
    )
    ALERT_POPUP = st.checkbox(
        "新预警弹窗告警", value=True,
        disabled=not WEATHER_AUTO_REFRESH, key="cfg_alert_popup",
    )

    if st.button("🔔 测试告警音", use_container_width=True,
                 disabled=not WEATHER_AUTO_REFRESH, key="btn_test_sound"):
        st.session_state["_force_play_sound"] = True

    if WEATHER_AUTO_REFRESH and not _HAS_AUTOREFRESH:
        st.warning("未安装 streamlit-autorefresh，自动刷新不可用。\n`pip install streamlit-autorefresh`")

    st.markdown("---")
    st.caption("💡 自动推送/自动刷新均依赖页面保持运行（浏览器标签页需保持打开）。")

# ==========================================
# 3.1 统一自动刷新调度
# ==========================================
_refresh_intervals = []
if WEATHER_AUTO_REFRESH:
    _refresh_intervals.append(int(WEATHER_REFRESH_MIN))
if AUTO_PUSH:
    _refresh_intervals.append(int(PUSH_INTERVAL_MIN))

if _refresh_intervals and _HAS_AUTOREFRESH:
    _tick_min = max(1, min(_refresh_intervals))
    st_autorefresh(interval=_tick_min * 60 * 1000, key="global_autorefresh_tick")
elif _HAS_AUTOREFRESH:
    _tick_min = None


def _weather_bucket():
    """按刷新间隔生成时间桶，桶号变化时自动让 st.cache_data 失效重取。"""
    if not WEATHER_AUTO_REFRESH:
        return 0
    interval_sec = max(60, int(WEATHER_REFRESH_MIN) * 60)
    return int(time.time() // interval_sec)


# ==========================================
# 4. 数据加载
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
# 5. 核心配置区
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

SEVERITY_ORDER = [
    ('红色', 0), ('Extreme', 0),
    ('橙色', 1), ('Severe', 1),
    ('黄色', 2), ('Moderate', 2),
    ('蓝色', 3), ('Minor', 3),
]


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
# 6. 工具函数
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


def _warn_field(warn, *keys, default=''):
    for k in keys:
        v = warn.get(k)
        if v:
            return str(v)
    return default


def _severity_rank(warn):
    sev = str(warn.get('severity', warn.get('level', 'Unknown'))).lower()
    for k, v in SEVERITY_ORDER:
        if k.lower() in sev:
            return v
    return 9


def _alert_key(city, warn):
    """预警唯一标识：城市 + 标题 + 等级 + 发布时间。"""
    return "|".join([
        str(city),
        _warn_field(warn, 'title', 'typeName', default=''),
        _warn_field(warn, 'severity', 'level', default=''),
        _warn_field(warn, 'pubTime', 'sent', default=''),
    ])


@st.cache_data(ttl=1800, show_spinner=False)
def get_weather_warnings_cached(lat, lon, bucket):
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


@st.cache_data(ttl=1800, show_spinner=False)
def get_current_weather_cached(lat, lon, bucket):
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
# 7. 板块数据处理
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
# 8. 城市代表坐标
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
# 9. 天气预警板块
# ==========================================
WEATHER_ALERT_STORE = {}   # {城市: [预警 dict, ...]}，供文字通报复用
WEATHER_ALERT_KEYS = {}    # {城市: set(本次成功获取的预警 key)}，供新预警检测


def render_global_weather_section():
    st.subheader("🌦️ 天气预警（全区域）")
    city_reps, err = build_city_representatives()
    if err:
        st.warning(err)
        return set()

    bucket = _weather_bucket()
    alert_cities = set()
    cols = st.columns(min(len(city_reps), 3))

    for i, (city, (lat, lon)) in enumerate(sorted(city_reps.items())):
        warnings, error = get_weather_warnings_cached(lat, lon, bucket)
        with cols[i % 3]:
            if error is not None:
                st.warning(f"⚠️ {city} 天气预警获取失败：{error}")
                continue
            if not warnings:
                st.success(f"✅ {city} 当前无天气预警")
                WEATHER_ALERT_KEYS[city] = set()
                continue

            alert_cities.add(city)
            WEATHER_ALERT_STORE[city] = list(warnings)
            WEATHER_ALERT_KEYS[city] = {_alert_key(city, w) for w in warnings}

            for warn in warnings[:3]:
                severity = warn.get('severity', warn.get('level', 'Unknown'))
                title = warn.get('title', warn.get('typeName', '天气预警'))
                text = warn.get('text', warn.get('description', ''))
                event = warn.get('typeName', warn.get('event', ''))
                sent = warn.get('pubTime', warn.get('sent', ''))
                instruction = warn.get('instruction', '')

                level_cn, color = SEVERITY_MAP.get(severity, ('未知', '#9B9B9B'))

                instruction_html = ""
                if instruction:
                    instruction_html = (
                        '<details style="margin-top:4px;">'
                        '<summary style="font-size:12px;color:#555;cursor:pointer;">展开防御指南</summary>'
                        f'<div style="font-size:12px;color:#555;margin-top:4px;">{instruction}</div>'
                        '</details>'
                    )

                card_html = (
                    f'<div style="border-left:5px solid {color};background:#F7F9FC;padding:10px 14px;'
                    'border-radius:6px;margin-bottom:8px;">'
                    f'<div style="font-weight:600;font-size:14px;color:{color};">{city} · {title}</div>'
                    f'<div style="font-size:12px;color:#666;margin:4px 0;">等级：{level_cn} · 类型：{event}</div>'
                    f'<div style="font-size:12px;color:#888;margin-top:6px;">发布时间：{sent}</div>'
                    '<details style="margin-top:6px;">'
                    '<summary style="font-size:12px;color:#555;cursor:pointer;">展开详细描述</summary>'
                    f'<div style="font-size:12px;color:#555;margin-top:4px;">{text}</div>'
                    '</details>'
                    f'{instruction_html}'
                    '</div>'
                )

                st.markdown(card_html, unsafe_allow_html=True)

    if alert_cities:
        st.warning(f"⚠️ 当前有天气预警的城市：{', '.join(sorted(alert_cities))}")
    else:
        st.success("✅ 所有涉及城市当前均无天气预警")
    return alert_cities


# ==========================================
# 9.1 新预警检测 + 告警提示音 + 弹窗
# ==========================================
def detect_new_alerts():
    """对比上次快照，返回本次新增的预警列表。"""
    seen = st.session_state.get("_weather_seen_keys")
    first_run = seen is None
    if seen is None:
        seen = {}

    new_items = []
    for city, keys in WEATHER_ALERT_KEYS.items():
        old_keys = seen.get(city, set())
        added = keys - old_keys
        if added and not first_run:
            for w in WEATHER_ALERT_STORE.get(city, []):
                if _alert_key(city, w) in added:
                    sev_raw = _warn_field(w, 'severity', 'level', default='Unknown')
                    level_cn, color = SEVERITY_MAP.get(sev_raw, ('未知', '#9B9B9B'))
                    new_items.append({
                        "city": city,
                        "title": _warn_field(w, 'title', 'typeName', default='天气预警'),
                        "level": level_cn,
                        "color": color,
                        "event": _warn_field(w, 'typeName', 'event', default='—'),
                        "sent": _warn_field(w, 'pubTime', 'sent', default='—'),
                        "text": _warn_field(w, 'text', 'description', default=''),
                    })
        seen[city] = keys

    st.session_state["_weather_seen_keys"] = seen
    return new_items


def play_alert_sound(volume=0.35):
    """注入 Web Audio 告警音（三短一长，类似警报器）。"""
    doc = f"""
    <script>
    (function() {{
      function beep() {{
        try {{
          var AC = window.AudioContext || window.webkitAudioContext;
          if (!AC) return;
          var ctx = new AC();
          if (ctx.state === 'suspended') {{ try {{ ctx.resume(); }} catch(e) {{}} }}
          var t0 = ctx.currentTime + 0.03;
          var seq = [880, 660, 880, 660, 1100];
          for (var i = 0; i < seq.length; i++) {{
            (function(i) {{
              var o = ctx.createOscillator();
              var g = ctx.createGain();
              o.type = 'square';
              o.frequency.value = seq[i];
              var st = t0 + i * 0.26;
              g.gain.setValueAtTime(0.0001, st);
              g.gain.exponentialRampToValueAtTime({volume}, st + 0.02);
              g.gain.exponentialRampToValueAtTime(0.0001, st + 0.22);
              o.connect(g); g.connect(ctx.destination);
              o.start(st); o.stop(st + 0.24);
            }})(i);
          }}
          setTimeout(function() {{ try {{ ctx.close(); }} catch(e) {{}} }}, 3000);
        }} catch (e) {{}}
      }}
      if (document.readyState === 'complete') beep();
      else window.addEventListener('load', beep);
      setTimeout(beep, 250);
    }})();
    </script>
    """
    components.html(doc, height=0)


def _alert_dialog_body(items, ts):
    st.markdown(
        "<div style='background:#FEF2F2;border-left:5px solid #D0021B;"
        "padding:10px 14px;border-radius:6px;margin-bottom:10px;'>"
        f"<div style='color:#B91C1C;font-weight:700;font-size:15px;'>"
        f"检测到 {len(items)} 条新天气预警</div>"
        f"<div style='color:#7F1D1D;font-size:12px;margin-top:4px;'>"
        f"发现时间：{ts:%Y-%m-%d %H:%M:%S}</div></div>",
        unsafe_allow_html=True,
    )
    for it in items:
        st.markdown(
            f"<div style='border-left:4px solid {it['color']};background:#F9FAFB;"
            "padding:8px 12px;border-radius:4px;margin-bottom:8px;'>"
            f"<div style='font-weight:600;color:{it['color']};font-size:13.5px;'>"
            f"{it['city']} · {it['title']}</div>"
            "<div style='font-size:12px;color:#6B7280;margin-top:3px;'>"
            f"等级：{it['level']} · 类型：{it['event']} · 发布时间：{it['sent']}</div>"
            "<div style='font-size:12px;color:#4B5563;margin-top:5px;'>"
            f"{it['text'][:200]}</div>"
            "</div>",
            unsafe_allow_html=True,
        )
    if st.button("✅ 我知道了", use_container_width=True, type="primary", key="ack_alert_dialog"):
        st.session_state.pop("_weather_new_alerts", None)
        st.session_state.pop("_weather_alert_ts", None)
        st.rerun()


if hasattr(st, "dialog"):
    @st.dialog("🚨 天气预警告警")
    def show_alert_dialog(items, ts):
        _alert_dialog_body(items, ts)
else:
    def show_alert_dialog(items, ts):
        st.error(f"🚨 检测到 {len(items)} 条新天气预警（{ts:%H:%M:%S}）")
        for it in items:
            st.markdown(f"- **{it['city']} · {it['title']}**（{it['level']}）")


# ==========================================
# 10. 当日城市天气板块
# ==========================================
def render_global_weather_now_section():
    st.subheader("🌤️ 当日城市天气")
    city_reps, err = build_city_representatives()
    if err:
        st.info(f"💡 {err}")
        return

    bucket = _weather_bucket()
    cols = st.columns(min(len(city_reps), 3))
    for i, (city, (lat, lon)) in enumerate(sorted(city_reps.items())):
        weather, error = get_current_weather_cached(lat, lon, bucket)
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

            now_html = (
                '<div style="background:linear-gradient(135deg,#E8F4FD 0%,#F7F9FC 100%);'
                'border:1px solid #D6E4F0;border-radius:10px;padding:16px 18px;margin-bottom:10px;">'
                '<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<div style="font-size:16px;font-weight:600;color:#2C5282;">🏙️ {city}</div>'
                f'<div style="font-size:32px;">{icon}</div>'
                '</div>'
                '<div style="font-size:28px;font-weight:700;color:#1A365D;margin:6px 0;">'
                f'{temp}<span style="font-size:16px;font-weight:400;">°C</span>'
                f'<span style="font-size:14px;font-weight:400;color:#4A5568;margin-left:8px;">{text}</span>'
                '</div>'
                '<div style="font-size:12px;color:#4A5568;margin-top:6px;line-height:1.8;">'
                f'🌡️ 体感温度：{feels}°C<br/>'
                f'💧 相对湿度：{humidity}%<br/>'
                f'🌬️ 风向风力：{wind_dir} {wind_scale}级'
                '</div>'
                '<div style="font-size:11px;color:#A0AEC0;margin-top:8px;text-align:right;">'
                f'更新时间：{time_str}'
                '</div>'
                '</div>'
            )
            st.markdown(now_html, unsafe_allow_html=True)


# ==========================================
# 11. 渲染两个独立板块（含新预警检测）
# ==========================================
with st.expander("🌦️ 天气预警板块", expanded=True):
    alert_cities_global = render_global_weather_section()

# ---- 新预警检测与告警 ----
_new_alerts = detect_new_alerts()

_force_sound = st.session_state.pop("_force_play_sound", False)

if _force_sound and ALERT_SOUND:
    play_alert_sound()

if _new_alerts:
    st.session_state["_weather_new_alerts"] = _new_alerts
    st.session_state["_weather_alert_ts"] = datetime.datetime.now()
    if ALERT_SOUND:
        play_alert_sound()
    cities_str = "、".join(sorted({a["city"] for a in _new_alerts}))
    st.toast(f"🚨 检测到 {len(_new_alerts)} 条新天气预警：{cities_str}", icon="🚨")

# 弹窗
_pending_alerts = st.session_state.get("_weather_new_alerts")
_alert_ts = st.session_state.get("_weather_alert_ts", datetime.datetime.now())
if _pending_alerts and ALERT_POPUP:
    show_alert_dialog(_pending_alerts, _alert_ts)

st.markdown("---")

with st.expander("🌤️ 当日城市天气板块", expanded=True):
    render_global_weather_now_section()

st.markdown("---")

# ==========================================
# 12. 异常天气文字版通报：函数定义（渲染位置在第 17 节末尾）
# ==========================================
def build_affected_projects_by_city():
    """返回 {城市: {板块: [项目名, ...]}}，只包含高风险项目。"""
    frames = []
    for cat, df, cfg in (
        ("防汛", hr_fangxun, TAB_CONFIG['防汛']),
        ("消防", hr_xiaofang, TAB_CONFIG['消防']),
        ("治安", hr_zhian, TAB_CONFIG['治安']),
    ):
        if df is None or len(df) == 0:
            continue
        point_col = cfg['point']
        if point_col not in df.columns:
            continue
        tmp = df.copy()
        tmp['_city'] = tmp.apply(extract_city, axis=1)
        tmp['_cat'] = cat
        tmp['_point'] = tmp[point_col].fillna('未知点位').astype(str)
        frames.append(tmp[['_city', '_cat', '_point']])

    result = {}
    if not frames:
        return result
    allp = pd.concat(frames, ignore_index=True)
    allp = allp[allp['_city'] != '其他']
    for (city, cat), g in allp.groupby(['_city', '_cat']):
        names = [n for n in g['_point'].tolist() if n and n.lower() != 'nan']
        if names:
            result.setdefault(city, {})[cat] = names
    return result


def build_weather_text_report(store, affected=None, brief=False, now=None):
    """生成异常天气文字版通报（纯文本）。"""
    now = now or datetime.datetime.now()
    sep = "─" * 32
    L = []
    L.append(sep)
    L.append("【异常天气预警通报】")
    L.append("发布单位：成都城市营业部")
    L.append(f"发布时间：{now:%Y-%m-%d %H:%M}")
    L.append(sep)
    L.append("")

    if not store:
        L.append("一、总体情况")
        L.append("截至目前，辖区内各城市均无生效中的天气预警。")
        L.append("")
        L.append("二、工作要求")
        L.append("请各片区保持关注天气变化，做好日常防范与值守工作。")
        L.append(sep)
        return "\n".join(L)

    items = []
    for city, warns in store.items():
        for w in warns:
            items.append((city, w))
    items.sort(key=lambda x: _severity_rank(x[1]))

    order_names = ['红色', '橙色', '黄色', '蓝色', '未知']
    counts = {k: 0 for k in order_names}
    for _, w in items:
        rank = _severity_rank(w)
        counts[order_names[rank] if rank < len(order_names) else '未知'] += 1
    cnt_str = "、".join(f"{k} {v} 条" for k, v in counts.items() if v > 0)

    L.append("一、总体情况")
    L.append(f"当前共 {len(store)} 个城市存在生效中的天气预警，合计 {len(items)} 条。")
    if cnt_str:
        L.append(f"等级分布：{cnt_str}。")
    L.append(f"涉及城市：{'、'.join(sorted(store.keys()))}。")
    L.append("")

    L.append("二、分城市预警详情")
    for city in sorted(store.keys()):
        warns = sorted(store[city], key=_severity_rank)
        L.append(f"■ {city}（{len(warns)} 条）")
        for i, w in enumerate(warns, 1):
            title = _warn_field(w, 'title', 'typeName', default='天气预警')
            sev_raw = _warn_field(w, 'severity', 'level', default='Unknown')
            level_cn, _ = SEVERITY_MAP.get(sev_raw, ('未知', ''))
            event = _warn_field(w, 'typeName', 'event', default='—')
            sent = _warn_field(w, 'pubTime', 'sent', default='—')
            text = _warn_field(w, 'text', 'description', default='')
            instruction = _warn_field(w, 'instruction', default='')

            L.append(f"  {i}.【{title}】等级：{level_cn} · 类型：{event}")
            L.append(f"     发布时间：{sent}")
            if not brief and text:
                L.append(f"     详情：{text}")
            if not brief and instruction:
                L.append(f"     防御指南：{instruction}")
        L.append("")

    alert_city_set = set(store.keys())
    if affected:
        hit = {c: v for c, v in affected.items() if c in alert_city_set}
        if hit:
            L.append("三、受影响高风险项目（按城市）")
            for city in sorted(hit.keys()):
                L.append(f"■ {city}")
                for cat in ("防汛", "消防", "治安"):
                    names = hit[city].get(cat)
                    if names:
                        L.append(f"    · {cat}（{len(names)} 个）：{'、'.join(names)}")
            L.append("")

    L.append("四、工作要求")
    L.append("1. 各片区收到通报后立即响应，明确值守责任人与应急联系方式。")
    L.append("2. 对红色、橙色预警区域的高风险项目开展现场巡查，重点排查排水、用电、围挡、高空坠物等隐患。")
    L.append("3. 落实物资储备（沙袋、水泵、应急照明等），确保随时可用。")
    L.append("4. 发现险情第一时间上报，保持信息畅通。")
    L.append(sep)
    return "\n".join(L)


def push_to_wecom_text(content, webhook_url, timeout=10):
    """推送纯文本到企业微信群机器人。返回 (成功?, 提示信息)。"""
    if not webhook_url or not str(webhook_url).strip():
        return False, "未配置企业微信机器人 Webhook 地址"
    webhook_url = str(webhook_url).strip()
    if not webhook_url.startswith("http"):
        return False, "Webhook 地址格式不正确"

    try:
        raw = content.encode("utf-8")
        limit = 1800
        if len(raw) > limit:
            content = raw[:limit].decode("utf-8", "ignore") + "\n…（内容过长已截断，完整通报请查看看板）"

        resp = requests.post(
            webhook_url,
            json={"msgtype": "text", "text": {"content": content}},
            timeout=timeout,
        )
        data = resp.json()
        if data.get("errcode") == 0:
            return True, "推送成功"
        return False, f"推送失败：errcode={data.get('errcode')} {data.get('errmsg', '')}"
    except Exception as e:
        return False, f"推送异常：{e}"


# ==========================================
# 13. 板块数据准备函数
# ==========================================
def build_category_detail(category_name, high_risk_df, config):
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
# 14. 微信联系人式清单
# ==========================================
def _esc(v):
    return html_module.escape(str(v) if v is not None else '')


def _build_contacts_doc(df, config, height=520, px_per_sec=22.0, paused=False):
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
# 15. 板块指标卡渲染函数（定义保留，渲染位置在第 17 节）
# ==========================================
def render_metric_card(title_emoji, title_text, total, n_city, n_street, accent_color="#D0021B"):
    card_html = (
        '<div style="border:1px solid #E5E7EB;border-radius:12px;padding:14px 16px;'
        'background:linear-gradient(135deg,#FFFFFF 0%,#F9FAFB 100%);height:100%;">'
        '<div style="font-size:15px;font-weight:600;color:#1A202C;margin-bottom:10px;">'
        f'{title_emoji} {title_text}板块'
        '</div>'
        '<div style="display:flex;justify-content:space-between;align-items:center;'
        'padding:8px 0;border-top:1px solid #F3F4F6;">'
        '<span style="color:#6B7280;font-size:12.5px;">🚨 高风险项目总数</span>'
        f'<span style="color:{accent_color};font-size:18px;font-weight:700;">{total}</span>'
        '</div>'
        '<div style="display:flex;justify-content:space-between;align-items:center;'
        'padding:8px 0;border-top:1px solid #F3F4F6;">'
        '<span style="color:#6B7280;font-size:12.5px;">🏢 高风险蝶城项目</span>'
        f'<span style="color:{accent_color};font-size:18px;font-weight:700;">{n_city}</span>'
        '</div>'
        '<div style="display:flex;justify-content:space-between;align-items:center;'
        'padding:8px 0;border-top:1px solid #F3F4F6;">'
        '<span style="color:#6B7280;font-size:12.5px;">🏘️ 高风险街区住宅</span>'
        f'<span style="color:{accent_color};font-size:18px;font-weight:700;">{n_street}</span>'
        '</div>'
        '</div>'
    )
    st.markdown(card_html, unsafe_allow_html=True)


# ==========================================
# 16. 三板块横向滚动清单
# ==========================================
def render_all_scrolling_boards(sections, height=520, default_speed=2.5):
    st.subheader("🚨 高风险点位清单")

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
# 17. 渲染顺序：滚动清单 → 指标卡 → 异常天气文字版通报
# ==========================================
detail_fangxun = build_category_detail("防汛", hr_fangxun, TAB_CONFIG['防汛'])
detail_xiaofang = build_category_detail("消防", hr_xiaofang, TAB_CONFIG['消防'])
detail_zhian = build_category_detail("治安", hr_zhian, TAB_CONFIG['治安'])

# ---- 17.1 高风险点位清单（三列滚动列表） ----
render_all_scrolling_boards(
    [
        ("防汛", detail_fangxun, TAB_CONFIG['防汛']),
        ("消防", detail_xiaofang, TAB_CONFIG['消防']),
        ("治安", detail_zhian, TAB_CONFIG['治安']),
    ],
    height=520,
    default_speed=2.5,
)

st.markdown("---")

# ---- 17.2 三大板块高风险指标卡 ----
st.subheader("📊 三大板块高风险指标")
metric_cols = st.columns(3)
with metric_cols[0]:
    render_metric_card("🌊", "防汛", len(hr_fangxun), n_city_fx, n_street_fx)
with metric_cols[1]:
    render_metric_card("🔥", "消防", len(hr_xiaofang), n_city_xf, n_street_xf)
with metric_cols[2]:
    render_metric_card("🚓", "治安", len(hr_zhian), n_city_za, n_street_za)

st.markdown("---")

# ---- 17.3 异常天气文字版通报（可推送企微） ----
affected_projects = build_affected_projects_by_city()
report_text = build_weather_text_report(WEATHER_ALERT_STORE, affected_projects)

with st.expander("📢 异常天气文字版通报（可推送企微）", expanded=True):
    if WEATHER_ALERT_STORE:
        total_alerts = sum(len(v) for v in WEATHER_ALERT_STORE.values())
        st.error(
            f"⚠️ 当前 **{len(WEATHER_ALERT_STORE)}** 个城市存在天气预警，"
            f"合计 **{total_alerts}** 条：{'、'.join(sorted(WEATHER_ALERT_STORE.keys()))}"
        )
    else:
        st.success("✅ 当前各城市均无生效中的天气预警，通报为常规版本。")

    st.text_area(
        "通报正文（点击右侧图标可直接复制）",
        report_text,
        height=340,
        key=f"weather_report_text_{abs(hash(report_text))}",
    )

    btn_c1, btn_c2, btn_c3 = st.columns([1.2, 1.2, 2])

    with btn_c1:
        st.download_button(
            "⬇️ 下载通报 (.txt)",
            data=report_text.encode("utf-8"),
            file_name=f"异常天气通报_{datetime.datetime.now():%Y%m%d_%H%M}.txt",
            mime="text/plain",
            use_container_width=True,
            key="download_weather_report_txt",
        )

    with btn_c2:
        manual_push = st.button("🚀 立即推送企微", use_container_width=True, type="primary")

    with btn_c3:
        if WECOM_WEBHOOK:
            st.caption("✅ 已配置 Webhook" + ("（自动推送已开启）" if AUTO_PUSH else ""))
        else:
            st.caption("⚠️ 未配置 Webhook，请到左侧边栏填写后即可推送")

    if manual_push:
        push_content = (
            build_weather_text_report(WEATHER_ALERT_STORE, affected_projects, brief=True)
            if BRIEF_PUSH else report_text
        )
        ok, msg = push_to_wecom_text(push_content, WECOM_WEBHOOK)
        st.session_state["_wx_last_push_time"] = datetime.datetime.now()
        if ok:
            st.success(f"✅ {msg}（{datetime.datetime.now():%H:%M:%S}）")
        else:
            st.error(f"❌ {msg}")

# ---- 17.4 定时自动推送（依赖统一 tick） ----
if AUTO_PUSH and _HAS_AUTOREFRESH:
    _now = datetime.datetime.now()
    _last = st.session_state.get("_wx_last_push_time")
    _due = (_last is None) or ((_now - _last).total_seconds() >= PUSH_INTERVAL_MIN * 60)

    if WECOM_WEBHOOK and _due:
        if ONLY_WHEN_ALERT and not WEATHER_ALERT_STORE:
            st.session_state["_wx_last_push_time"] = _now
            st.toast("ℹ️ 当前无预警，本次自动推送已跳过", icon="ℹ️")
        else:
            _push_content = (
                build_weather_text_report(WEATHER_ALERT_STORE, affected_projects, brief=True)
                if BRIEF_PUSH else report_text
            )
            _ok, _msg = push_to_wecom_text(_push_content, WECOM_WEBHOOK)
            st.session_state["_wx_last_push_time"] = _now
            if _ok:
                st.toast(f"✅ 企微自动推送成功（{_now:%H:%M:%S}）", icon="✅")
            else:
                st.toast(f"❌ 企微自动推送失败：{_msg}", icon="⚠️")

if AUTO_PUSH and WECOM_WEBHOOK:
    _last_show = st.session_state.get("_wx_last_push_time")
    st.caption(
        f"⏱️ 自动推送已启用，间隔 {PUSH_INTERVAL_MIN} 分钟 · "
        f"上次处理时间：{_last_show.strftime('%Y-%m-%d %H:%M:%S') if _last_show else '尚未推送'}"
    )

if WEATHER_AUTO_REFRESH and _HAS_AUTOREFRESH:
    st.caption(
        f"⏱️ 天气自动刷新已启用，间隔 {WEATHER_REFRESH_MIN} 分钟 · "
        f"上次刷新：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}"
    )

# ==========================================
# 18. 一键导出完整报告
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
# 19. 原始数据展开查看
# ==========================================
st.markdown("---")
with st.expander("点击查看原始数据表格"):
    st.dataframe(filtered_df, use_container_width=True)
# -*- coding: utf-8 -*-
"""
航班行程安排助手（Excel 计划版）
上传航段数据 Excel → 选择飞机 → 填需求 → 输出该飞机完整行程 + 新增段
"""

import streamlit as st
import re
import io
from math import radians, sin, cos, asin, sqrt
from datetime import datetime, timedelta, date
from openpyxl import load_workbook

st.set_page_config(page_title="航班行程安排助手", page_icon="✈️", layout="wide")
st.title("✈️ 航班行程安排助手")
st.caption("上传航段数据 Excel → 选择飞机 → 填需求 → 自动排出完整行程")


# ================================================================
# 城市 → 四字码
# ================================================================
CITY_TO_ICAO = {
    "澳门": "VMMC", "香港": "VHHH", "台北松山": "RCSS", "台北桃园": "RCTP",
    "北京大兴": "ZBAD", "北京首都": "ZBAA",
    "上海虹桥": "ZSSS", "上海浦东": "ZSPD",
    "成都双流": "ZUUU", "成都天府": "ZUTF",
    "深圳宝安": "ZGSZ", "广州白云": "ZGGG",
    "宁波栎社": "ZSNB", "温州龙湾": "ZSWZ",
    "厦门高崎": "ZSAM", "泉州晋江": "ZSQZ",
    "杭州萧山": "ZSHC", "南京禄口": "ZSNJ",
    "天津滨海": "ZBTJ", "沈阳桃仙": "ZYTX",
    "乌鲁木齐天山": "ZWWW", "重庆江北": "ZUCK",
    "西宁曹家堡": "ZLXN", "鄂尔多斯伊金霍洛": "ZBDS",
    "丽江三义": "ZPLJ", "威海大水泊": "ZSWH",
    "长沙黄花": "ZGHA", "嘉兴南湖": "ZSJX",
    "宜昌三峡": "ZHYC", "济宁大安": "ZSJG",
    "福州长乐": "ZSFZ", "青岛胶东": "ZSQD",
    "武汉天河": "ZHHH", "海口美兰": "ZJHK",
    "三亚凤凰": "ZJSY", "珠海金湾": "ZGSD",
    "大连周水子": "ZYTL", "合肥新桥": "ZSOF",
    "南昌昌北": "ZSCN", "贵阳龙洞堡": "ZUGY",
    "桂林两江": "ZGKL", "哈尔滨太平": "ZYHB",
    "长春龙嘉": "ZYCC", "台中清泉岗": "RCMQ",
    "越南岘港": "VVDN", "柬埔寨金边 德崇": "VDTI", "柬埔寨金边": "VDTI",
    "老挝万象": "VLVT", "瓦岱": "VLVT",
    "日本东京 羽田": "RJTT", "日本东京 成田": "RJAA",
    "日本大阪 关西": "RJBB", "日本那霸": "ROAH",
    "韩国济州岛": "RKPC", "韩国首尔": "RKSI",
    "新加坡实里达": "WSSL", "新加坡樟宜": "WSSS",
    "泰国曼谷 廊曼": "VTBD", "泰国普吉": "VTSP",
    "泰国清莱": "VTCT", "菲律宾马尼拉": "RPLL",
    "印尼万鸦老": "WAMM", "印尼韦达港": "WAEH",
    "印尼雅加达 哈达": "WIII", "印尼龙目岛普拉亚": "WADL",
    "澳大利亚悉尼 金斯福德": "YSSY",
    "澳大利亚墨尔本": "YMML",
    "澳大利亚朗塞斯顿": "YMLT",
    "加拿大蒙特利尔 特鲁多": "CYUL",
    "加拿大温哥华": "CYVR",
    "美国贝德福德": "KBED", "美国达拉斯": "KDFW",
    "美国费城": "KPHL", "美国旧金山": "KSFO",
    "美国加利福尼亚州圣迭戈": "KSDM",
    "马尔代夫马法鲁岛": "VRDA",
    "捷克布拉格 鲁济涅": "LKPR",
    "法国巴黎 布尔歇": "LFPB",
    "西班牙马德里 巴拉哈斯": "LEMD",
    "罗马": "LIRA", "埃及卢克索": "HELX",
    "埃及阿布辛伯勒": "HEBL", "埃及古尔代盖": "HEGN",
    "埃及开罗": "HECA", "文莱斯里巴加湾": "WBSB",
    "厄瓜多尔基多": "SEQM", "秘鲁库斯科": "SPZO",
    "玻利维亚乌尤尼": "SLUY", "智利伊基克": "SCDA",
    "智利卡拉马": "SCCF", "秘鲁利马": "SPJC",
    "巴哈马拿骚": "MYNN",
}


def get_icao(city):
    if not city:
        return None
    city = str(city).strip()
    if city in CITY_TO_ICAO:
        return CITY_TO_ICAO[city]
    for k, v in CITY_TO_ICAO.items():
        if k in city or city in k:
            return v
    return None


# ================================================================
# 机场坐标
# ================================================================
AIRPORT_COORDS = {
    "ZBAD": (39.509, 116.411), "ZBAA": (40.079, 116.585),
    "ZSSS": (31.197, 121.336), "ZSPD": (31.143, 121.805),
    "ZUUU": (30.578, 103.947), "ZUTF": (30.312, 104.441),
    "ZGSZ": (22.639, 113.811), "ZGGG": (23.392, 113.298),
    "ZSNB": (29.826, 121.462), "ZSWZ": (27.912, 120.852),
    "ZSAM": (24.544, 118.128), "ZSQZ": (24.796, 118.590),
    "ZSHC": (30.229, 120.434), "ZSNJ": (31.742, 118.862),
    "ZBTJ": (39.124, 117.346), "ZYTX": (41.639, 123.483),
    "ZWWW": (43.907, 87.474), "ZUCK": (29.719, 106.641),
    "ZLXN": (36.528, 102.043), "ZBDS": (39.494, 109.859),
    "ZPLJ": (26.680, 100.246), "ZSWH": (37.187, 122.229),
    "ZGHA": (28.189, 113.219), "ZSJX": (30.701, 120.665),
    "ZHYC": (30.552, 111.479), "ZSJG": (35.496, 116.740),
    "ZSFZ": (25.935, 119.663), "ZSQD": (36.361, 120.086),
    "ZHHH": (30.784, 114.208), "ZJHK": (19.935, 110.459),
    "ZJSY": (18.303, 109.412), "ZGSD": (22.006, 113.376),
    "ZYTL": (38.966, 121.538), "ZSOF": (31.780, 116.976),
    "ZSCN": (28.865, 115.900), "ZUGY": (26.539, 106.801),
    "ZGKL": (25.218, 110.039), "ZYHB": (45.623, 126.250),
    "ZYCC": (43.996, 125.685),
    "VMMC": (22.149, 113.592), "VHHH": (22.309, 113.915),
    "RCSS": (25.069, 121.552), "RCTP": (25.077, 121.233),
    "RCMQ": (24.264, 120.621),
    "VVDN": (16.044, 108.199), "VDTI": (11.547, 104.844),
    "VLVT": (17.988, 102.563),
    "RJTT": (35.549, 139.779), "RJAA": (35.765, 140.386),
    "RJBB": (34.434, 135.232), "ROAH": (26.196, 127.646),
    "RKPC": (33.512, 126.493), "RKSI": (37.469, 126.451),
    "WSSL": (1.417, 103.868), "WSSS": (1.357, 103.988),
    "VTBD": (13.912, 100.607), "VTSP": (8.113, 98.317),
    "VTCT": (19.953, 99.883), "RPLL": (14.508, 121.020),
    "WAMM": (1.549, 124.926), "WAEH": (-0.836, 127.786),
    "WIII": (-6.125, 106.656), "WADL": (-8.757, 116.276),
    "YSSY": (-33.946, 151.177), "YMML": (-37.673, 144.843),
    "YMLT": (-41.545, 147.214),
    "CYUL": (45.470, -73.741), "CYVR": (49.196, -123.182),
    "KBED": (42.470, -71.289), "KDFW": (32.897, -97.038),
    "KPHL": (39.872, -75.241), "KSFO": (37.619, -122.375),
    "KSDM": (32.572, -116.980),
    "VRDA": (4.191, 73.500),
    "LKPR": (50.101, 14.260), "LFPB": (48.969, 2.441),
    "LEMD": (40.472, -3.561), "LIRA": (41.799, 12.594),
    "HELX": (25.671, 32.707), "HEBL": (22.376, 31.611),
    "HEGN": (27.178, 33.799), "HECA": (30.111, 31.413),
    "WBSB": (4.944, 114.928),
    "SEQM": (-0.129, -78.358), "SPZO": (-13.536, -71.939),
    "SLUY": (-20.442, -66.848), "SCDA": (-20.535, -70.181),
    "SCCF": (-22.498, -68.904), "SPJC": (-12.022, -77.114),
    "MYNN": (25.039, -77.466),
}


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat / 2) ** 2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon / 2) ** 2
    return 2 * R * asin(sqrt(a))


def estimate_flight_minutes(from_icao, to_icao):
    if not from_icao or not to_icao:
        return None
    if from_icao == to_icao:
        return 0
    if from_icao not in AIRPORT_COORDS or to_icao not in AIRPORT_COORDS:
        return None
    lat1, lon1 = AIRPORT_COORDS[from_icao]
    lat2, lon2 = AIRPORT_COORDS[to_icao]
    d = haversine_km(lat1, lon1, lat2, lon2)
    return round(d / 800 * 60 + 25)


# ================================================================
# 时间工具
# ================================================================
def time_to_min(t):
    """接受 datetime.time / datetime.datetime / 'HH:MM' 字符串"""
    if t is None:
        return None
    if isinstance(t, datetime):
        return t.hour * 60 + t.minute
    try:
        if hasattr(t, "hour") and hasattr(t, "minute"):
            return t.hour * 60 + t.minute
    except Exception:
        pass
    m = re.match(r"^(\d{1,2}):(\d{2})$", str(t).strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def min_to_time_str(m):
    m = m % 1440
    return f"{m // 60:02d}:{m % 60:02d}"


def min_to_dur_str(m):
    h, mi = divmod(int(m), 60)
    if h and mi:
        return f"{h}h{mi:02d}m"
    elif h:
        return f"{h}h"
    else:
        return f"{mi}m"


def plan_minutes(flight_minutes):
    total = flight_minutes + 15
    return ((total + 4) // 5) * 5


def is_domestic(icao):
    return bool(icao) and str(icao).startswith("Z")


def transit_minutes(icao):
    return 90 if is_domestic(icao) else 120


def to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


# ================================================================
# 解析 Excel 计划
# ================================================================
def load_excel_plan(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb["航段(北京时)"] if "航段(北京时)" in wb.sheetnames else wb.active

    rows = []
    for r in range(3, ws.max_row + 1):
        reg = ws.cell(r, 3).value
        if not reg:
            continue

        reg_str = str(reg).strip().upper()

        dep_date = to_date(ws.cell(r, 7).value)
        arr_date = to_date(ws.cell(r, 15).value)

        dep_icao = str(ws.cell(r, 11).value or "").strip().upper()
        dep_city = str(ws.cell(r, 12).value or "").strip()
        arr_icao = str(ws.cell(r, 13).value or "").strip().upper()
        arr_city = str(ws.cell(r, 14).value or "").strip()

        # 城市可能有多个机场名，取第一个
        dep_city_clean = dep_city.split()[0] if dep_city else ""
        arr_city_clean = arr_city.split()[0] if arr_city else ""

        dep_min = time_to_min(ws.cell(r, 8).value)
        arr_min = time_to_min(ws.cell(r, 16).value)

        if dep_min is None or arr_min is None or not dep_date:
            continue

        # 若到达时间 < 出发时间，视为跨天（到达日期 +1）
        arr_dt_date = arr_date or dep_date
        if arr_date is None and arr_min < dep_min:
            arr_dt_date = dep_date + timedelta(days=1)

        rows.append({
            "reg": reg_str,
            "use": str(ws.cell(r, 4).value or "").strip(),
            "dep_date": dep_date,
            "dep_min": dep_min,
            "dep_icao": dep_icao,
            "dep_city": dep_city_clean,
            "arr_date": arr_dt_date,
            "arr_min": arr_min,
            "arr_icao": arr_icao,
            "arr_city": arr_city_clean,
        })

    return rows


# ================================================================
# 解析需求
# ================================================================
def parse_request(text):
    clean_text = text

    target_date = None
    m = re.search(r"(\d{1,2})\s*[.\-/月]\s*(\d{1,2})\s*[日号]?", clean_text)
    if m:
        mon, day = int(m.group(1)), int(m.group(2))
        year = datetime.now().year
        try:
            target_date = date(year, mon, day)
        except Exception:
            pass

    route = None
    m = re.search(
        r"([\u4e00-\u9fff][\u4e00-\u9fff\s]{0,20}?)\s*[-–—到至]\s*"
        r"([\u4e00-\u9fff][\u4e00-\u9fff\s]{0,20})",
        clean_text,
    )
    if m:
        dep_c = m.group(1).strip()
        arr_c = m.group(2).strip()
        dep_c = re.sub(r"^\d+[.\-/月]\d+\s*[日号]?\s*", "", dep_c).strip()
        if dep_c and arr_c:
            route = (dep_c, arr_c)

    flight_min = None
    for c in re.findall(r"(?<!\d)(\d{4})(?!\d)", clean_text):
        hh, mm = int(c[:2]), int(c[2:])
        if hh < 24 and mm < 60:
            flight_min = hh * 60 + mm
            break

    if flight_min is None:
        m = re.search(
            r"(\d{1,2})\s*(?:h|H|小时)\s*(?:(\d{1,2})\s*(?:m|M|分钟?))?",
            clean_text,
        )
        if m:
            hh = int(m.group(1))
            mm = int(m.group(2)) if m.group(2) else 0
            if hh < 24 and mm < 60:
                flight_min = hh * 60 + mm

    if flight_min is None:
        m = re.search(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)", clean_text)
        if m:
            hh, mm = int(m.group(1)), int(m.group(2))
            if hh < 24 and mm < 60:
                flight_min = hh * 60 + mm

    return {
        "date": target_date,
        "route": route,
        "flight_min": flight_min,
    }


# ================================================================
# 分析某架飞机
# ================================================================
DUTY_MAX_MIN = 14 * 60
FLIGHT_MAX_MIN = 10 * 60
REST_MIN = 10 * 60


def analyze_aircraft(flights_all, reg, target_date, dep_city, arr_city,
                     flight_min, transit, plan_min):
    dep_icao_req = get_icao(dep_city)
    arr_icao_req = get_icao(arr_city)
    if not dep_icao_req:
        return {"error": f"❌ 无法识别出发城市：{dep_city}"}
    if not arr_icao_req:
        return {"error": f"❌ 无法识别到达城市：{arr_city}"}

    reg_flights = [f for f in flights_all if f["reg"] == reg]
    reg_flights.sort(key=lambda x: (x["dep_date"], x["dep_min"]))

    before = [f for f in reg_flights if f["dep_date"] <= target_date]
    day_flights = [f for f in reg_flights if f["dep_date"] == target_date]

    if not before:
        return {"error": f"❌ 飞机 {reg} 在 {target_date.strftime('%m月%d日')} 之前没有任何航段。"}

    # ---- 判断是否需要调机 ----
    ferry_info = None
    if day_flights:
        last_of_day = day_flights[-1]
        if last_of_day["arr_icao"] == dep_icao_req:
            ferry_info = None
        else:
            ferry_from_icao = last_of_day["arr_icao"]
            ferry_from_city = last_of_day["arr_city"]
            ferry_min = estimate_flight_minutes(ferry_from_icao, dep_icao_req)
            if ferry_min is None:
                return {"error": f"❌ 缺少坐标数据，无法估算 {ferry_from_icao} → {dep_icao_req} 调机时间。"}
            ferry_dep_min = last_of_day["arr_min"] + transit_minutes(ferry_from_icao)
            ferry_plan_min = plan_minutes(ferry_min)
            ferry_arr_min = ferry_dep_min + ferry_plan_min
            ferry_info = {
                "from_city": ferry_from_city,
                "from_icao": ferry_from_icao,
                "ferry_min": ferry_min,
                "ferry_plan_min": ferry_plan_min,
                "dep_min": ferry_dep_min,
                "arr_min": ferry_arr_min,
                "based_on": last_of_day,
            }
    else:
        last_seg = before[-1]
        last_city_icao = last_seg["arr_icao"]
        if last_city_icao != dep_icao_req:
            ferry_from_city = last_seg["arr_city"]
            ferry_from_icao = last_city_icao
            ferry_min = estimate_flight_minutes(ferry_from_icao, dep_icao_req)
            if ferry_min is None:
                return {"error": f"❌ 缺少坐标数据，无法估算 {ferry_from_icao} → {dep_icao_req} 调机时间。"}
            ferry_dep_min = 8 * 60
            ferry_plan_min = plan_minutes(ferry_min)
            ferry_arr_min = ferry_dep_min + ferry_plan_min
            ferry_info = {
                "from_city": ferry_from_city,
                "from_icao": ferry_from_icao,
                "ferry_min": ferry_min,
                "ferry_plan_min": ferry_plan_min,
                "dep_min": ferry_dep_min,
                "arr_min": ferry_arr_min,
                "based_on": last_seg,
            }
        else:
            ferry_dep_min = 8 * 60
            ferry_arr_min = ferry_dep_min
            ferry_info = {
                "from_city": dep_city,
                "from_icao": dep_icao_req,
                "ferry_min": 0,
                "ferry_plan_min": 0,
                "dep_min": ferry_dep_min,
                "arr_min": ferry_arr_min,
                "based_on": last_seg,
            }

    # ---- 主段起飞 ----
    if ferry_info:
        main_dep_min = ferry_info["arr_min"] + transit
    else:
        main_dep_min = day_flights[-1]["arr_min"] + transit

    main_arr_min = main_dep_min + plan_min

    # ---- 值勤时间 ----
    if day_flights:
        duty_start = day_flights[0]["dep_min"] - 120
    elif ferry_info:
        duty_start = ferry_info["dep_min"] - 120
    else:
        duty_start = main_dep_min - 120

    duty_end = main_arr_min + 60
    duty_total = duty_end - duty_start

    # ---- 飞行总时长 ----
    old_flight = 0
    for f in day_flights:
        d = f["dep_min"]
        a = f["arr_min"]
        if a < d:
            a += 1440
        old_flight += (a - d)

    ferry_flight = ferry_info["ferry_min"] if ferry_info else 0
    total_flight = old_flight + ferry_flight + flight_min

    # ---- 前日休息 ----
    prev_day = target_date - timedelta(days=1)
    rest_ok = True
    rest_note = ""
    prev_day_flights = [f for f in reg_flights if f["dep_date"] == prev_day]
    if prev_day_flights:
        prev_last_arr = prev_day_flights[-1]["arr_min"]
        prev_duty_end = prev_last_arr + 60
        rest_hours = (duty_start - prev_duty_end) / 60
        if rest_hours < 10:
            rest_ok = False
            rest_note = (
                f"前一日值勤结束 {min_to_time_str(prev_duty_end)}，"
                f"当日值勤开始 {min_to_time_str(duty_start)}，"
                f"休息仅 {rest_hours:.1f}h < 10h"
            )

    # ---- 组装航段 ----
    segments = []
    for f in reg_flights:
        segments.append({
            "date": f["dep_date"],
            "dep_city": f["dep_city"],
            "dep_icao": f["dep_icao"],
            "dep_time": min_to_time_str(f["dep_min"]),
            "arr_time": min_to_time_str(f["arr_min"]),
            "arr_city": f["arr_city"],
            "arr_icao": f["arr_icao"],
            "tag": "",       # 无标记
        })

    if ferry_info and ferry_info["ferry_min"] > 0:
        segments.append({
            "date": target_date,
            "dep_city": ferry_info["from_city"],
            "dep_icao": ferry_info["from_icao"],
            "dep_time": min_to_time_str(ferry_info["dep_min"]),
            "arr_time": min_to_time_str(ferry_info["arr_min"]),
            "arr_city": dep_city,
            "arr_icao": dep_icao_req,
            "tag": "调机",
        })

    segments.append({
        "date": target_date,
        "dep_city": dep_city,
        "dep_icao": dep_icao_req,
        "dep_time": min_to_time_str(main_dep_min),
        "arr_time": min_to_time_str(main_arr_min),
        "arr_city": arr_city,
        "arr_icao": arr_icao_req,
        "tag": "新增",
    })

    return {
        "reg": reg,
        "segments": segments,
        "day_flights": day_flights,
        "duty_start": duty_start,
        "duty_end": duty_end,
        "duty_total": duty_total,
        "old_flight": old_flight,
        "ferry_info": ferry_info,
        "transit": transit,
        "main_dep": main_dep_min,
        "main_arr": main_arr_min,
        "new_duty_end": duty_end,
        "new_duty_total": duty_total,
        "new_flight_total": total_flight,
        "ok_duty": duty_total <= DUTY_MAX_MIN,
        "ok_flight": total_flight <= FLIGHT_MAX_MIN,
        "rest_ok": rest_ok,
        "rest_note": rest_note,
        "plan_min": plan_min,
        "flight_min": flight_min,
    }


# ================================================================
# UI
# ================================================================

uploaded = st.file_uploader(
    "① 上传航段数据导出 (.xlsx)", type=["xlsx"], key="xlsx_plan"
)

flights = []
all_regs = []
if uploaded is not None:
    try:
        flights = load_excel_plan(uploaded.getvalue())
        all_regs = sorted(set(f["reg"] for f in flights))
    except Exception as e:
        st.error(f"❌ 无法解析 Excel：{e}")

col_a, col_b = st.columns([1, 2])

with col_a:
    st.subheader("② 选择飞机")
    if all_regs:
        selected_reg = st.selectbox(
            "选择飞机", options=all_regs,
            label_visibility="collapsed", key="reg_select",
        )
    else:
        st.info("请先上传 Excel")
        selected_reg = None

with col_b:
    st.subheader("③ 潜在行程请求")
    request_text = st.text_area(
        "request", height=130, label_visibility="collapsed",
        placeholder="例如：\n10.4 澳门-老挝万象 0158\n起飞时间可以适当根据我们调配",
        key="request_input",
    )

st.write("")
run = st.button("🚀 分析并推荐", type="primary", use_container_width=True)


# ================================================================
# 主流程
# ================================================================
if run:
    if not flights:
        st.error("请先上传航段数据 Excel")
        st.stop()
    if not selected_reg:
        st.error("请选择飞机")
        st.stop()
    if not request_text.strip():
        st.error("请填写潜在行程请求")
        st.stop()

    request = parse_request(request_text)
    missing = []
    if not request["date"]:
        missing.append("**日期**（如 `10.4` 或 `10月4日`）")
    if not request["route"]:
        missing.append("**航线**（如 `澳门-老挝万象`）")
    if not request["flight_min"]:
        missing.append("**飞行时间**（如 `0158`、`1h58m`）")
    if missing:
        st.error("❌ 需求描述中缺少：")
        for item in missing:
            st.markdown(f"- {item}")
        st.stop()

    target_date = request["date"]
    dep_city, arr_city = request["route"]
    flight_min = request["flight_min"]
    dep_icao = get_icao(dep_city)
    arr_icao = get_icao(arr_city)
    plan_min = plan_minutes(flight_min)
    transit = transit_minutes(dep_icao) if dep_icao else 120

    st.subheader("📋 请求摘要")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("日期", str(target_date))
    c2.metric("飞机", selected_reg)
    c3.metric("航线", f"{dep_icao or dep_city} → {arr_icao or arr_city}")
    c4.metric("飞行时间", min_to_dur_str(flight_min))
    c5.metric("计划时间", min_to_dur_str(plan_min))
    st.caption(
        f"过站时间：**{min_to_dur_str(transit)}**"
        f"（{'国内' if is_domestic(dep_icao) else '国际'}）"
    )

    st.markdown("---")

    res = analyze_aircraft(
        flights, selected_reg, target_date,
        dep_city, arr_city, flight_min, transit, plan_min,
    )

    if res.get("error"):
        st.error(res["error"])
        st.stop()

    ok_all = res["ok_duty"] and res["ok_flight"] and res["rest_ok"]
    head_icon = "✅" if ok_all else "⚠️"
    st.subheader(f"{head_icon} {selected_reg} 新增后的完整行程")

    def _compact_line(seg):
        dep_hm = seg["dep_time"].replace(":", "")
        arr_hm = seg["arr_time"].replace(":", "")
        tag = f"  {seg['tag']}" if seg.get("tag") else ""
        return (
            f"{seg['date'].day}号 {seg['dep_city']}{dep_hm} "
            f"{arr_hm}{seg['arr_city']}{tag}"
        )

    compact_lines = [selected_reg]
    for seg in res["segments"]:
        compact_lines.append(_compact_line(seg))
    st.code("\n".join(compact_lines), language=None)

    if res["ferry_info"] and res["ferry_info"]["ferry_min"] > 0:
        fi = res["ferry_info"]
        st.markdown(
            f"**调机段：** `{min_to_time_str(fi['dep_min'])} - "
            f"{min_to_time_str(fi['arr_min'])}`  "
            f"{fi['from_city']} → {dep_city}  "
            f"（估算飞行 {min_to_dur_str(fi['ferry_min'])}）"
        )
        st.caption(
            f"飞机原计划 {fi['based_on']['dep_date'].strftime('%m月%d日')} "
            f"{min_to_time_str(fi['based_on']['arr_min'])} 到达 {fi['from_city']}"
        )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f"**当日现有值勤：** {min_to_time_str(res['duty_start'])} → "
            f"{min_to_time_str(res['duty_start'] + res['duty_total'])} "
            f"（{min_to_dur_str(res['duty_total'])}）"
        )
        st.markdown(f"**当日现有飞行：** {min_to_dur_str(res['old_flight'])}")
    with c2:
        st.markdown(
            f"**推荐主段：** `{min_to_time_str(res['main_dep'])} - "
            f"{min_to_time_str(res['main_arr'])}`  "
            f"{dep_city} → {arr_city}"
        )
        st.markdown(
            f"**新增后值勤：** {min_to_time_str(res['duty_start'])} → "
            f"{min_to_time_str(res['new_duty_end'])} "
            f"（{min_to_dur_str(res['new_duty_total'])}）"
        )
        st.markdown(f"**新增后飞行：** {min_to_dur_str(res['new_flight_total'])}")

    checks = [
        f"{'✅' if res['ok_duty'] else '❌'} 值勤 ≤ 14h"
        f"（{min_to_dur_str(res['new_duty_total'])}）",
        f"{'✅' if res['ok_flight'] else '❌'} 飞行 ≤ 10h"
        f"（{min_to_dur_str(res['new_flight_total'])}）",
        f"{'✅' if res['rest_ok'] else '❌'} 前日休息 ≥ 10h"
        + (f"（{res['rest_note']}）" if res["rest_note"] else ""),
    ]
    for c in checks:
        st.markdown(f"- {c}")

    st.markdown("---")
    st.subheader("📄 可复制方案")
    st.caption("点右上角复制按钮，直接粘贴到 Jetops / 邮件 / 微信")
    st.code("\n".join(compact_lines), language=None)

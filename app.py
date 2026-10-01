# -*- coding: utf-8 -*-
"""
航班行程安排助手（含自动调机估算）
"""

import streamlit as st
import re
from math import radians, sin, cos, asin, sqrt
from datetime import datetime, timedelta, date

st.set_page_config(page_title="航班行程安排助手", page_icon="✈️", layout="wide")
st.title("✈️ 航班行程安排助手")
st.caption("输入现有航班计划和潜在行程，自动分析并推荐最佳安排（无飞机在出发城市时自动估算调机）")


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
    city = city.strip()
    if city in CITY_TO_ICAO:
        return CITY_TO_ICAO[city]
    for k, v in CITY_TO_ICAO.items():
        if k in city or city in k:
            return v
    return None


# ================================================================
# 机场经纬度（用于调机时间估算）
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
    """估算两个机场间的飞行时间（分钟）。巡航 ~800 km/h，+ 25 min 起降。"""
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
def time_str_to_min(t):
    m = re.match(r"^(\d{1,2}):(\d{2})$", str(t).strip())
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


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
    return bool(icao) and icao.startswith("Z")


def transit_minutes(icao):
    return 90 if is_domestic(icao) else 120


# ================================================================
# 解析航班计划
# ================================================================
FLIGHT_HEADER_RE = re.compile(
    r"^([A-Z0-9]+)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(\s*\+1)?$"
)


def parse_schedule(text):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    flights = []
    i = 0
    pending_f = False

    while i < len(lines):
        line = lines[i]

        if line.upper() == "F":
            pending_f = True
            i += 1
            continue
        if line.upper() == "TBA":
            i += 1
            continue

        m = FLIGHT_HEADER_RE.match(line)
        if m:
            reg = m.group(1).upper()
            dep_time = m.group(2)
            arr_time = m.group(3)
            plus1 = bool(m.group(4))

            if i + 1 < len(lines):
                cm = re.match(r"^(.+?)\s+-\s+(.+)$", lines[i + 1])
                if cm:
                    dep_city = cm.group(1).strip()
                    arr_city = cm.group(2).strip()

                    crew = []
                    step = 2
                    if i + 2 < len(lines):
                        cl = lines[i + 2].replace(" ", "")
                        if re.match(r"^[A-Z0-9,]+$", cl):
                            crew = [x for x in cl.split(",") if x]
                            step = 3

                    flights.append({
                        "reg": reg,
                        "dep_time": dep_time,
                        "arr_time": arr_time,
                        "dep_city": dep_city,
                        "arr_city": arr_city,
                        "crew": crew,
                        "is_ferry": pending_f,
                        "plus1": plus1,
                    })
                    pending_f = False
                    i += step
                    continue
        i += 1
    return flights


def assign_dates(flights, start_date):
    current = start_date
    prev_min = None
    for f in flights:
        dm = time_str_to_min(f["dep_time"])
        if dm is None:
            continue
        if prev_min is not None and dm < prev_min:
            current += timedelta(days=1)
        f["date"] = current
        prev_min = dm
    return flights


# ================================================================
# 解析潜在行程请求
# ================================================================
def _is_schedule_line(s):
    if not s:
        return True
    if s in ("F", "TBA"):
        return True
    if re.match(r"^[A-Z0-9]+\s+\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", s):
        return True
    if len(s) > 3 and re.match(r"^[A-Z0-9,]+$", s):
        return True
    return False


def parse_request(text):
    raw_lines = text.splitlines()
    keep = []
    skip_city = False
    for ln in raw_lines:
        s = ln.strip()
        if not s:
            continue
        if _is_schedule_line(s):
            if re.match(r"^[A-Z0-9]+\s+\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", s):
                skip_city = True
            continue
        if skip_city and re.match(r"^[\u4e00-\u9fff\s]+\s*-\s*[\u4e00-\u9fff\s]+$", s):
            skip_city = False
            continue
        skip_city = False
        keep.append(s)

    clean_text = "\n".join(keep)
    no_content = not clean_text.strip()

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

    reg = None
    m = re.search(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]{3,7})(?![A-Za-z0-9])", clean_text)
    if m:
        reg = m.group(1)

    return {
        "date": target_date, "route": route, "flight_min": flight_min,
        "reg": reg, "raw": clean_text, "no_content": no_content,
    }


# ================================================================
# 分析
# ================================================================
DUTY_MAX_MIN = 14 * 60
FLIGHT_MAX_MIN = 10 * 60
REST_MIN = 10 * 60


def analyze(flights, request):
    if request.get("no_content"):
        return {"errors": ["NO_CONTENT"]}

    missing = []
    if not request["date"]:
        missing.append("**日期**（如 `10.4` 或 `10月4日`）")
    if not request["route"]:
        missing.append("**航线**（如 `澳门-老挝万象`）")
    if not request["flight_min"]:
        missing.append("**飞行时间**（如 `0158`、`1h58m`）")
    if missing:
        return {"errors": ["MISSING_FIELDS", missing]}

    target_date = request["date"]
    dep_city, arr_city = request["route"]
    dep_icao = get_icao(dep_city)
    arr_icao = get_icao(arr_city)

    if not dep_icao:
        return {"errors": [f"❌ 无法识别出发城市：{dep_city}"]}
    if not arr_icao:
        return {"errors": [f"❌ 无法识别到达城市：{arr_city}"]}

    flight_min = request["flight_min"]
    plan_min = plan_minutes(flight_min)
    transit = transit_minutes(dep_icao)

    by_date_reg = {}
    for f in flights:
        by_date_reg.setdefault((f["date"], f["reg"]), []).append(f)

    all_regs = sorted(set(f["reg"] for f in flights))
    if request["reg"]:
        all_regs = [r for r in all_regs if r == request["reg"]] or all_regs

    results = []

    for reg in all_regs:
        reg_flights = [f for f in flights if f["reg"] == reg]
        reg_flights.sort(key=lambda x: (x["date"], time_str_to_min(x["dep_time"]) or 0))

        before = [f for f in reg_flights if f["date"] <= target_date]
        day_flights = [f for f in reg_flights if f["date"] == target_date]
        day_flights.sort(key=lambda x: time_str_to_min(x["dep_time"]) or 0)

        # ============ 情况 A：当天在 dep_city ============
        if day_flights:
            day_last = day_flights[-1]
            last_arr_icao = get_icao(day_last["arr_city"])
            if last_arr_icao == dep_icao:
                first_dep = time_str_to_min(day_flights[0]["dep_time"])
                last_arr = time_str_to_min(day_last["arr_time"])
                duty_start = first_dep - 120
                duty_end = last_arr + 60

                old_flight = 0
                for f in day_flights:
                    d = time_str_to_min(f["dep_time"])
                    a = time_str_to_min(f["arr_time"])
                    if d is not None and a is not None:
                        old_flight += (a - d) if a >= d else (a + 1440 - d)

                main_dep = last_arr + transit
                main_arr = main_dep + plan_min
                new_duty_end = main_arr + 60
                new_duty_total = new_duty_end - duty_start
                new_flight_total = old_flight + flight_min

                results.append({
                    "reg": reg,
                    "day_flights": day_flights,
                    "crew": day_flights[0].get("crew", []),
                    "duty_start": duty_start, "duty_end": duty_end,
                    "duty_total": duty_end - duty_start,
                    "old_flight": old_flight,
                    "ferry_info": None,
                    "transit": transit,
                    "main_dep": main_dep, "main_arr": main_arr,
                    "new_duty_end": new_duty_end,
                    "new_duty_total": new_duty_total,
                    "new_flight_total": new_flight_total,
                    "ok_duty": new_duty_total <= DUTY_MAX_MIN,
                    "ok_flight": new_flight_total <= FLIGHT_MAX_MIN,
                    "rest_ok": True, "rest_note": "",
                    "plan_min": plan_min, "flight_min": flight_min,
                })
                continue

        # ============ 情况 B：需要调机 ============
        if not before:
            continue

        last_seg = before[-1]
        last_city_icao = get_icao(last_seg["arr_city"])
        if last_city_icao == dep_icao:
            # 就是前一天到的，日期偏差问题不大，按当天处理
            if last_seg["date"] != target_date:
                # 前一天到 dep_city，可以第二天一早出发
                ferry_dep_min = 8 * 60  # 默认 08:00 出发（可调）
                ferry_min = 0
                ferry_plan_min = 0
                ferry_arr_min = ferry_dep_min
            else:
                continue
        else:
            ferry_min = estimate_flight_minutes(last_city_icao, dep_icao)
            if ferry_min is None or ferry_min <= 0:
                continue

            # 调机段起飞时间
            if last_seg["date"] == target_date:
                last_arr_min = time_str_to_min(last_seg["arr_time"])
                transit_before_ferry = transit_minutes(last_city_icao)
                ferry_dep_min = last_arr_min + transit_before_ferry
            else:
                ferry_dep_min = 8 * 60  # 前一天或更早到的，默认次日 08:00

            ferry_plan_min = plan_minutes(ferry_min)
            ferry_arr_min = ferry_dep_min + ferry_plan_min

        # 调机到达后，过站 + 主段
        main_dep_min = ferry_arr_min + transit_minutes(dep_icao)
        main_arr_min = main_dep_min + plan_min

        # 值勤时间
        if last_seg["date"] == target_date:
            # 当天已有航段，从当天第一段起飞前 2h 开始
            first_dep = time_str_to_min(before[0]["dep_time"]) if before else ferry_dep_min
            for f in before:
                if f["date"] == target_date:
                    first_dep = time_str_to_min(f["dep_time"])
                    break
            duty_start = first_dep - 120
        else:
            duty_start = ferry_dep_min - 120

        duty_end = main_arr_min + 60
        duty_total = duty_end - duty_start

        # 飞行总时长
        old_flight = 0
        for f in before:
            if f["date"] == target_date:
                d = time_str_to_min(f["dep_time"])
                a = time_str_to_min(f["arr_time"])
                if d is not None and a is not None:
                    old_flight += (a - d) if a >= d else (a + 1440 - d)

        total_flight = old_flight + (ferry_min or 0) + flight_min

        # 前日休息
        prev_day = target_date - timedelta(days=1)
        rest_ok = True
        rest_note = ""
        for (pd, preg), prev_flights in by_date_reg.items():
            if pd == prev_day and preg == reg:
                prev_sorted = sorted(prev_flights, key=lambda x: time_str_to_min(x["dep_time"]))
                prev_last_arr = time_str_to_min(prev_sorted[-1]["arr_time"])
                prev_duty_end = prev_last_arr + 60
                rest_hours = (duty_start - prev_duty_end) / 60
                if rest_hours < 10:
                    rest_ok = False
                    rest_note = (
                        f"前一日值勤结束 {min_to_time_str(prev_duty_end)}，"
                        f"当日值勤开始 {min_to_time_str(duty_start)}，"
                        f"休息仅 {rest_hours:.1f}h < 10h"
                    )
                break

        results.append({
            "reg": reg,
            "day_flights": day_flights,
            "crew": (day_flights[0].get("crew") if day_flights else (before[0].get("crew") if before else [])),
            "duty_start": duty_start, "duty_end": duty_end,
            "duty_total": duty_end - duty_start,
            "old_flight": old_flight,
            "ferry_info": {
                "from_city": last_seg["arr_city"],
                "from_icao": last_city_icao,
                "from_seg": last_seg,
                "ferry_min": ferry_min,
                "ferry_plan_min": ferry_plan_min,
                "dep_min": ferry_dep_min,
                "arr_min": ferry_arr_min,
            },
            "transit": transit,
            "main_dep": main_dep_min, "main_arr": main_arr_min,
            "new_duty_end": duty_end,
            "new_duty_total": duty_total,
            "new_flight_total": total_flight,
            "ok_duty": duty_total <= DUTY_MAX_MIN,
            "ok_flight": total_flight <= FLIGHT_MAX_MIN,
            "rest_ok": rest_ok, "rest_note": rest_note,
            "plan_min": plan_min, "flight_min": flight_min,
        })

    if not results:
        return {"errors": [
            f"⚠️ 无法为 {target_date.strftime('%m月%d日')} 的 {dep_city}（{dep_icao}）出发行程找到可行方案。",
            "可能原因：",
            "  · 没有飞机在计划内能及时赶到 " + dep_city,
            "  · 缺少 " + dep_city + " 的坐标数据，无法估算调机时间",
        ]}

    # 排序：可行的在前
    results.sort(key=lambda r: (
        0 if (r["ok_duty"] and r["ok_flight"] and r["rest_ok"]) else 1,
        r["new_duty_total"],
    ))

    return {
        "errors": [],
        "date": target_date,
        "route": (dep_city, arr_city),
        "dep_icao": dep_icao,
        "arr_icao": arr_icao,
        "flight_min": flight_min,
        "plan_min": plan_min,
        "transit": transit,
        "results": results,
    }


# ================================================================
# UI
# ================================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("① 现有航班计划")
    st.caption("粘贴当前航班计划（含机组代号），F 表示调机")
    schedule_text = st.text_area(
        "schedule", height=420, label_visibility="collapsed",
        placeholder=(
            "F\nB652Q 13:20 - 16:20\n澳门 - 北京大兴\n"
            "P002,P068,P046,C036,M021\n\n..."
        ),
        key="schedule_input",
    )

with col2:
    st.subheader("② 潜在行程请求")
    st.caption("只写需求：日期 + 航线 + 飞行时间（如 `10.4 澳门-老挝万象 0158`）")
    request_text = st.text_area(
        "request", height=420, label_visibility="collapsed",
        placeholder=(
            "例如：\n辛苦评估一下共享租赁计划：\n"
            "10.4 澳门-老挝万象 0158\n还是明天 MLLIN 这波客人的行程"
        ),
        key="request_input",
    )

st.markdown("---")

col_d, col_b = st.columns([1, 3])
with col_d:
    start_date = st.date_input("计划第一天日期", value=date.today())
with col_b:
    st.write("")
    st.write("")
    run = st.button("🚀 分析并推荐", type="primary", use_container_width=True)

if run:
    if not schedule_text.strip():
        st.error("请粘贴航班计划")
    elif not request_text.strip():
        st.error("请填写潜在行程请求")
    else:
        flights = parse_schedule(schedule_text)
        if not flights:
            st.error("❌ 无法解析航班计划，请检查格式")
        else:
            flights = assign_dates(flights, start_date)
            request = parse_request(request_text)
            result = analyze(flights, request)

            if result.get("errors"):
                errs = result["errors"]
                if errs and errs[0] == "NO_CONTENT":
                    st.error("❌ 右侧请求框没有有效需求描述")
                    st.markdown(
                        "看起来你把**航班计划**粘到了右边。请分开：\n\n"
                        "📌 **左侧**：完整航班计划\n"
                        "📌 **右侧**：需求描述，例如 `10.4 澳门-老挝万象 0158`"
                    )
                elif errs and errs[0] == "MISSING_FIELDS":
                    st.error("❌ 需求描述中缺少以下信息：")
                    for item in errs[1]:
                        st.markdown(f"- {item}")
                else:
                    for e in errs:
                        st.warning(e)
            else:
                st.subheader("📋 请求摘要")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("日期", str(result["date"]))
                c2.metric("航线", f"{result['dep_icao']} → {result['arr_icao']}")
                c3.metric("飞行时间", min_to_dur_str(result["flight_min"]))
                c4.metric("计划时间", min_to_dur_str(result["plan_min"]))
                st.caption(
                    f"过站时间：**{min_to_dur_str(result['transit'])}**"
                    f"（{'国内' if is_domestic(result['dep_icao']) else '国际'}）"
                )

                st.markdown("---")
                st.subheader("✈️ 候选方案")

                out_lines = []

                def _seg_sort_key(seg):
                    return (seg["date"], time_str_to_min(seg["dep_time"]) or 0)

                def _compact_line(seg):
                    dep_hm = seg["dep_time"].replace(":", "")
                    arr_hm = seg["arr_time"].replace(":", "")
                    if seg.get("is_ferry_leg"):
                        tag = "  ← 调机"
                    elif seg.get("is_new"):
                        tag = "  ← 推荐新增"
                    else:
                        tag = "  调机" if seg.get("is_ferry") else ""
                    return (
                        f"{seg['date'].day}号 {seg['dep_city']}{dep_hm} "
                        f"{arr_hm}{seg['arr_city']}{tag}"
                    )

                for r in result["results"]:
                    reg = r["reg"]
                    duty_start_str = min_to_time_str(r["duty_start"])
                    duty_end_str = min_to_time_str(r["duty_end"])
                    main_dep_str = min_to_time_str(r["main_dep"])
                    main_arr_str = min_to_time_str(r["main_arr"])

                    ok_all = r["ok_duty"] and r["ok_flight"] and r["rest_ok"]
                    head_icon = "✅" if ok_all else "⚠️"

                    # 组装完整航段列表（该飞机全部）
                    all_of_reg = [f for f in flights if f["reg"] == reg]
                    all_of_reg.sort(key=_seg_sort_key)

                    all_segments = []
                    for f in all_of_reg:
                        all_segments.append({
                            "date": f["date"], "dep_city": f["dep_city"],
                            "dep_time": f["dep_time"], "arr_time": f["arr_time"],
                            "arr_city": f["arr_city"], "is_ferry": f["is_ferry"],
                            "is_new": False, "is_ferry_leg": False,
                        })

                    # 插入调机段
                    if r["ferry_info"]:
                        fi = r["ferry_info"]
                        all_segments.append({
                            "date": result["date"],
                            "dep_city": fi["from_city"],
                            "dep_time": min_to_time_str(fi["dep_min"]),
                            "arr_time": min_to_time_str(fi["arr_min"]),
                            "arr_city": result["route"][0],
                            "is_ferry": False, "is_new": False,
                            "is_ferry_leg": True,
                        })

                    # 插入主段
                    all_segments.append({
                        "date": result["date"],
                        "dep_city": result["route"][0],
                        "dep_time": main_dep_str,
                        "arr_time": main_arr_str,
                        "arr_city": result["route"][1],
                        "is_ferry": False, "is_new": True, "is_ferry_leg": False,
                    })
                    all_segments.sort(key=_seg_sort_key)

                    with st.container():
                        st.markdown(f"### {head_icon} {reg}")

                        st.markdown("**该飞机完整行程（紧凑格式）：**")
                        compact_lines = [reg]
                        for seg in all_segments:
                            compact_lines.append(_compact_line(seg))
                        st.code("\n".join(compact_lines), language=None)

                        # 调机信息
                        if r["ferry_info"]:
                            fi = r["ferry_info"]
                            last_seg = fi["from_seg"]
                            st.markdown(
                                f"**调机段：** `{min_to_time_str(fi['dep_min'])} - "
                                f"{min_to_time_str(fi['arr_min'])}`  "
                                f"{fi['from_city']} → {result['route'][0]}  "
                                f"（估算飞行 {min_to_dur_str(fi['ferry_min'])}）"
                            )
                            st.caption(
                                f"飞机当前位于 {fi['from_city']}"
                                f"（{last_seg['date'].strftime('%m月%d日')} "
                                f"{last_seg['arr_time']} 到达）"
                            )

                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(
                                f"**当日现有值勤：** {duty_start_str} → {duty_end_str} "
                                f"（{min_to_dur_str(r['duty_total'])}）"
                            )
                            st.markdown(f"**当日现有飞行：** {min_to_dur_str(r['old_flight'])}")
                            crew_str = ", ".join(r["crew"]) if r["crew"] else "—"
                            st.markdown(f"**机组：** `{crew_str}`")
                        with c2:
                            st.markdown(
                                f"**推荐主段：** `{main_dep_str} - {main_arr_str}`  "
                                f"{result['route'][0]} → {result['route'][1]}"
                            )
                            st.markdown(
                                f"**新增后值勤：** {duty_start_str} → "
                                f"{min_to_time_str(r['new_duty_end'])} "
                                f"（{min_to_dur_str(r['new_duty_total'])}）"
                            )
                            st.markdown(
                                f"**新增后飞行：** {min_to_dur_str(r['new_flight_total'])}"
                            )

                        checks = [
                            f"{'✅' if r['ok_duty'] else '❌'} 值勤 ≤ 14h"
                            f"（{min_to_dur_str(r['new_duty_total'])}）",
                            f"{'✅' if r['ok_flight'] else '❌'} 飞行 ≤ 10h"
                            f"（{min_to_dur_str(r['new_flight_total'])}）",
                            f"{'✅' if r['rest_ok'] else '❌'} 前日休息 ≥ 10h"
                            + (f"（{r['rest_note']}）" if r["rest_note"] else ""),
                        ]
                        for c in checks:
                            st.markdown(f"- {c}")

                        st.markdown("---")

                    out_lines.extend(compact_lines)
                    out_lines.append("")

                st.subheader("📄 可复制方案")
                st.caption("点右上角复制按钮，直接粘贴到 Jetops / 邮件 / 微信")
                st.code("\n".join(out_lines).rstrip(), language=None)

# -*- coding: utf-8 -*-
"""
航班行程安排助手
输入现有航班计划 + 潜在行程请求，自动分析机组值勤约束并推荐最佳安排。
"""

import streamlit as st
import re
from datetime import datetime, timedelta, date

st.set_page_config(page_title="航班行程安排助手", page_icon="✈️", layout="wide")
st.title("✈️ 航班行程安排助手")
st.caption("输入现有航班计划和潜在行程，自动分析并推荐最佳安排")


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
    "越南岘港": "VVDN", "柬埔寨金边 德崇": "VDTI",
    "柬埔寨金边": "VDTI",
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
    """城市名 → 四字码"""
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
# 时间/时长工具
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
    """飞行时间 + 15分钟滑行，向上取整至5分钟"""
    total = flight_minutes + 15
    return ((total + 4) // 5) * 5


def is_domestic(icao):
    return bool(icao) and icao.startswith("Z")


def transit_minutes(dep_icao):
    """过站时间：国内 1.5h，国际 2h"""
    return 90 if is_domestic(dep_icao) else 120


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
    """按出现顺序分配日期；出发时间回退时视为新的一天"""
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
    """判断一行是否像航班计划（用于请求框里过滤误粘的计划）"""
    if not s:
        return True
    if s in ("F", "TBA"):
        return True
    # 航班头：XXX HH:MM - HH:MM
    if re.match(r"^[A-Z0-9]+\s+\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", s):
        return True
    # 机组行：全是 P/C/M/W/数字/逗号，且较长
    if len(s) > 3 and re.match(r"^[A-Z0-9,]+$", s):
        return True
    return False


def parse_request(text):
    # ---- 先过滤掉误粘的航班计划行 ----
    raw_lines = text.splitlines()
    keep = []
    skip_city = False
    for ln in raw_lines:
        s = ln.strip()
        if not s:
            continue
        if _is_schedule_line(s):
            # 航班头后跟一行城市，也要跳过
            if re.match(r"^[A-Z0-9]+\s+\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", s):
                skip_city = True
            continue
        if skip_city and re.match(r"^[\u4e00-\u9fff\s]+\s*-\s*[\u4e00-\u9fff\s]+$", s):
            skip_city = False
            continue
        skip_city = False
        keep.append(s)

    clean_text = "\n".join(keep)

    # ---- 日期 ----
    target_date = None
    m = re.search(r"(\d{1,2})\s*[.\-/月]\s*(\d{1,2})", clean_text)
    if m:
        mon, day = int(m.group(1)), int(m.group(2))
        year = datetime.now().year
        try:
            target_date = date(year, mon, day)
        except Exception:
            pass

    # ---- 航线 ----
    route = None
    m = re.search(
        r"([\u4e00-\u9fff][\u4e00-\u9fff\s]{1,30}?)\s*[-–—到至]\s*"
        r"([\u4e00-\u9fff][\u4e00-\u9fff\s]{1,30})",
        clean_text,
    )
    if m:
        route = (m.group(1).strip(), m.group(2).strip())

    # ---- 飞行时间：HHMM ----
    flight_min = None
    candidates = re.findall(r"(?<!\d)(\d{4})(?!\d)", clean_text)
    for c in candidates:
        hh, mm = int(c[:2]), int(c[2:])
        if hh < 24 and mm < 60:
            flight_min = hh * 60 + mm
            break

    # ---- 注册号（可选）----
    reg = None
    m = re.search(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]{3,7})(?![A-Za-z0-9])", clean_text)
    if m:
        reg = m.group(1)

    return {
        "date": target_date,
        "route": route,
        "flight_min": flight_min,
        "reg": reg,
        "raw": clean_text,
    }

# ================================================================
# 分析
# ================================================================
DUTY_MAX_MIN = 14 * 60      # 每日最大值勤
FLIGHT_MAX_MIN = 10 * 60    # 每日最大飞行
REST_MIN = 10 * 60          # 最小休息时间（分钟）


def analyze(flights, request):
    errors = []
    if not request["date"]:
        errors.append("❌ 无法从请求中识别日期，请写清楚，例如 `10.4` 或 `10月4日`")
    if not request["route"]:
        errors.append("❌ 无法识别航线，请写清楚，例如 `澳门-老挝万象`")
    if not request["flight_min"]:
        errors.append("❌ 无法识别飞行时间，请写 4 位数字，例如 `0158` 表示 1h58m")
    if errors:
        return {"errors": errors}

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
        key = (f["date"], f["reg"])
        by_date_reg.setdefault(key, []).append(f)

    candidates = []
    for (d, reg), day_flights in by_date_reg.items():
        if d != target_date:
            continue
        day_sorted = sorted(day_flights, key=lambda x: time_str_to_min(x["dep_time"]))
        last_arr_icao = get_icao(day_sorted[-1]["arr_city"])
        if last_arr_icao == dep_icao:
            candidates.append((reg, day_sorted))

    if request["reg"]:
        specific = [c for c in candidates if c[0] == request["reg"]]
        if specific:
            candidates = specific

    if not candidates:
        return {"errors": [
            f"⚠️ 在 {target_date} 未找到当日最后一段降落在 {dep_city}（{dep_icao}）的飞机。",
            "请检查：",
            "  · 请求日期是否与计划中的某一天匹配？",
            "  · 计划中是否有飞机当日降落在该城市？",
        ]}

    results = []
    for reg, day_sorted in candidates:
        first_dep = time_str_to_min(day_sorted[0]["dep_time"])
        last_arr = time_str_to_min(day_sorted[-1]["arr_time"])

        duty_start = first_dep - 120
        duty_end = last_arr + 60
        duty_total = duty_end - duty_start

        old_flight = 0
        for f in day_sorted:
            d = time_str_to_min(f["dep_time"])
            a = time_str_to_min(f["arr_time"])
            if d is not None and a is not None:
                old_flight += (a - d) if a >= d else (a + 1440 - d)

        earliest_dep = last_arr + transit
        earliest_arr = earliest_dep + plan_min

        new_duty_end = earliest_arr + 60
        new_duty_total = new_duty_end - duty_start
        new_flight_total = old_flight + flight_min

        latest_arr = duty_start + DUTY_MAX_MIN - 60
        latest_dep = latest_arr - plan_min

        ok_duty = new_duty_total <= DUTY_MAX_MIN
        ok_flight = new_flight_total <= FLIGHT_MAX_MIN

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
            "day_flights": day_sorted,
            "crew": day_sorted[0].get("crew", []),
            "duty_start": duty_start,
            "duty_end": duty_end,
            "duty_total": duty_total,
            "old_flight": old_flight,
            "transit": transit,
            "earliest_dep": earliest_dep,
            "earliest_arr": earliest_arr,
            "latest_dep": latest_dep,
            "new_duty_end": new_duty_end,
            "new_duty_total": new_duty_total,
            "new_flight_total": new_flight_total,
            "ok_duty": ok_duty,
            "ok_flight": ok_flight,
            "rest_ok": rest_ok,
            "rest_note": rest_note,
            "plan_min": plan_min,
            "flight_min": flight_min,
        })

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
        "schedule",
        height=420,
        label_visibility="collapsed",
        placeholder=(
            "F\n"
            "B652Q 13:20 - 16:20\n"
            "澳门 - 北京大兴\n"
            "P002,P068,P046,C036,M021\n"
            "\n"
            "B652Q 18:20 - 21:00\n"
            "北京大兴 - 成都双流\n"
            "P002,PJZ001,C036,M021\n"
            "..."
        ),
        key="schedule_input",
    )

with col2:
    st.subheader("② 潜在行程请求")
    st.caption("描述新行程需求：日期 + 航线 + 飞行时间（4 位 HHMM）")
    request_text = st.text_area(
        "request",
        height=420,
        label_visibility="collapsed",
        placeholder=(
            "例如：\n"
            "辛苦评估一下共享租赁计划：\n"
            "10.4 澳门-老挝万象 0158\n"
            "还是明天 MLLIN 这波客人的行程\n"
            "这个起飞时间可以适当根据我们调配"
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
            st.error("❌ 无法解析航班计划，请检查格式（每段应为：注册号 起飞-落地 / 城市-城市 / 机组）")
        else:
            flights = assign_dates(flights, start_date)
            request = parse_request(request_text)
            result = analyze(flights, request)

            if result.get("errors"):
                for e in result["errors"]:
                    st.warning(e)
            else:
                st.subheader("📋 请求摘要")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("日期", str(result["date"]))
                c2.metric("航线", f"{result['dep_icao']} → {result['arr_icao']}")
                c3.metric("飞行时间", min_to_dur_str(result["flight_min"]))
                c4.metric("计划时间", min_to_dur_str(result["plan_min"]))

                st.caption(
                    f"过站时间要求：**{min_to_dur_str(result['transit'])}**"
                    f"（{'国内' if is_domestic(result['dep_icao']) else '国际'}）"
                )

                st.markdown("---")
                st.subheader("✈️ 候选飞机分析")

                out_lines = []

                def _seg_sort_key(seg):
                    return (seg["date"], time_str_to_min(seg["dep_time"]) or 0)

                def _compact_line(seg):
                    dep_hm = seg["dep_time"].replace(":", "")
                    arr_hm = seg["arr_time"].replace(":", "")
                    ferry = "  调机" if seg.get("is_ferry") else ""
                    new_tag = "  ← 推荐新增" if seg.get("is_new") else ""
                    return (
                        f"{seg['date'].day}号 {seg['dep_city']}{dep_hm} "
                        f"{arr_hm}{seg['arr_city']}{ferry}{new_tag}"
                    )

                for r in result["results"]:
                    reg = r["reg"]
                    duty_start_str = min_to_time_str(r["duty_start"])
                    duty_end_str = min_to_time_str(r["duty_end"])
                    new_duty_end_str = min_to_time_str(r["new_duty_end"])
                    rec_dep_str = min_to_time_str(r["earliest_dep"])
                    rec_arr_str = min_to_time_str(r["earliest_arr"])

                    ok_all = r["ok_duty"] and r["ok_flight"] and r["rest_ok"]
                    head_icon = "✅" if ok_all else "⚠️"

                    # ---- 该飞机在计划中所有航段 ----
                    all_of_reg = [f for f in flights if f["reg"] == reg]
                    all_of_reg.sort(key=_seg_sort_key)

                    all_segments = []
                    for f in all_of_reg:
                        all_segments.append({
                            "date": f["date"],
                            "dep_city": f["dep_city"],
                            "dep_time": f["dep_time"],
                            "arr_time": f["arr_time"],
                            "arr_city": f["arr_city"],
                            "is_ferry": f["is_ferry"],
                            "is_new": False,
                        })

                    # ---- 插入新增段 ----
                    new_arr_min = r["earliest_arr"]
                    new_arr_date = result["date"]
                    if new_arr_min >= 1440:
                        new_arr_date = result["date"] + timedelta(days=1)

                    all_segments.append({
                        "date": result["date"],
                        "dep_city": result["route"][0],
                        "dep_time": rec_dep_str,
                        "arr_time": min_to_time_str(new_arr_min),
                        "arr_city": result["route"][1],
                        "is_ferry": False,
                        "is_new": True,
                    })
                    all_segments.sort(key=_seg_sort_key)

                    with st.container():
                        st.markdown(f"### {head_icon} {reg}")

                        st.markdown("**该飞机完整行程（紧凑格式）：**")
                        compact_lines = [reg]
                        for seg in all_segments:
                            compact_lines.append(_compact_line(seg))
                        st.code("\n".join(compact_lines), language=None)

                        crew_str = ", ".join(r["crew"]) if r["crew"] else "—"

                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(
                                f"**当日现有值勤：** {duty_start_str} → {duty_end_str} "
                                f"（{min_to_dur_str(r['duty_total'])}）"
                            )
                            st.markdown(
                                f"**当日现有飞行：** {min_to_dur_str(r['old_flight'])}"
                            )
                            st.markdown(f"**机组：** `{crew_str}`")
                        with c2:
                            st.markdown(
                                f"**新增后值勤：** {duty_start_str} → {new_duty_end_str} "
                                f"（{min_to_dur_str(r['new_duty_total'])}）"
                            )
                            st.markdown(
                                f"**新增后飞行：** {min_to_dur_str(r['new_flight_total'])}"
                            )
                            st.markdown(
                                f"**推荐新增：** `{rec_dep_str} - {rec_arr_str}`  "
                                f"{result['route'][0]} → {result['route'][1]}"
                            )

                        checks = []
                        checks.append(
                            f"{'✅' if r['ok_duty'] else '❌'} 值勤 ≤ 14h"
                            f"（{min_to_dur_str(r['new_duty_total'])}）"
                        )
                        checks.append(
                            f"{'✅' if r['ok_flight'] else '❌'} 飞行 ≤ 10h"
                            f"（{min_to_dur_str(r['new_flight_total'])}）"
                        )
                        checks.append(
                            f"{'✅' if r['rest_ok'] else '❌'} 前日休息 ≥ 10h"
                            + (f"（{r['rest_note']}）" if r["rest_note"] else "")
                        )
                        for c in checks:
                            st.markdown(f"- {c}")

                        if r["latest_dep"] > r["earliest_dep"]:
                            with st.expander("备选出发时间（值勤范围内）"):
                                t = r["earliest_dep"]
                                cnt = 0
                                alt_lines = []
                                while t <= r["latest_dep"] and cnt < 10:
                                    a = t + r["plan_min"]
                                    alt_lines.append(
                                        f"{min_to_time_str(t)} - {min_to_time_str(a)}"
                                    )
                                    t += 30
                                    cnt += 1
                                st.code("\n".join(alt_lines), language=None)

                        st.markdown("---")

                    out_lines.extend(compact_lines)
                    out_lines.append("")

                st.subheader("📄 可复制方案")
                st.caption("点右上角复制按钮，直接粘贴到 Jetops / 邮件 / 微信")
                st.code("\n".join(out_lines).rstrip(), language=None)

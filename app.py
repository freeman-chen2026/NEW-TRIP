def analyze_aircraft(flights_all, reg, target_date, dep_city, arr_city,
                     flight_min, transit, plan_min, forced_main_dep=None,
                     merge_ferry=False):
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

    # ---------- 调机合并 ----------
    merged_info = None
    ferry_info = None

    if merge_ferry:
        last_pax = None
        later_ferry = []
        for f in day_flights:
            if is_ferry_use(f.get("use", "")):
                if last_pax is not None:
                    later_ferry.append(f)
            else:
                last_pax = f
                later_ferry = []

        if last_pax and later_ferry:
            final_ferry = later_ferry[-1]
            if final_ferry["arr_icao"] != dep_icao_req:
                ferry_from_icao = last_pax["arr_icao"]
                ferry_from_city = last_pax["arr_city"]
                last_pax_arr_abs = (
                    (last_pax["arr_date"] - target_date).days * 1440
                    + last_pax["arr_min"]
                )

                if ferry_from_icao == dep_icao_req:
                    ferry_min = 0
                    ferry_plan_min = 0
                    ferry_dep_min = last_pax_arr_abs
                    ferry_arr_min = ferry_dep_min
                else:
                    ferry_min = estimate_flight_minutes(ferry_from_icao, dep_icao_req)
                    if ferry_min is not None:
                        ferry_plan_min = plan_minutes(ferry_min)
                        ferry_dep_min = last_pax_arr_abs + transit_minutes(ferry_from_icao)
                        ferry_arr_min = ferry_dep_min + ferry_plan_min

                if ferry_min is not None:
                    ferry_info = {
                        "from_city": ferry_from_city,
                        "from_icao": ferry_from_icao,
                        "ferry_min": ferry_min,
                        "ferry_plan_min": ferry_plan_min,
                        "dep_min": ferry_dep_min,
                        "arr_min": ferry_arr_min,
                        "based_on": last_pax,
                    }
                    merged_info = {
                        "skipped": later_ferry,
                        "last_pax": last_pax,
                        "return_to_icao": final_ferry["arr_icao"],
                        "return_to_city": final_ferry["arr_city"],
                    }

    # ---------- 方案 A（默认调机）----------
    if ferry_info is None:
        if day_flights:
            last_of_day = day_flights[-1]
            if last_of_day["arr_icao"] == dep_icao_req:
                ferry_info = None
            else:
                ferry_from_icao = last_of_day["arr_icao"]
                ferry_from_city = last_of_day["arr_city"]
                ferry_min = estimate_flight_minutes(ferry_from_icao, dep_icao_req)
                if ferry_min is None:
                    return {"error": f"❌ 缺少坐标，无法估算 {ferry_from_icao} → {dep_icao_req} 调机时间。"}
                last_arr_abs = (
                    (last_of_day["arr_date"] - target_date).days * 1440
                    + last_of_day["arr_min"]
                )
                ferry_dep_min = last_arr_abs + transit_minutes(ferry_from_icao)
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
                    return {"error": f"❌ 缺少坐标，无法估算 {ferry_from_icao} → {dep_icao_req} 调机时间。"}
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

    # ---------- 主段 ----------
    if ferry_info is not None:
        earliest_main_dep = ferry_info["arr_min"] + transit
    else:
        last_arr_abs = (
            (day_flights[-1]["arr_date"] - target_date).days * 1440
            + day_flights[-1]["arr_min"]
        )
        earliest_main_dep = last_arr_abs + transit

    if forced_main_dep is not None:
        if forced_main_dep < earliest_main_dep:
            return {"error": (
                f"起飞 {min_to_time_str(forced_main_dep)} "
                f"早于最早可起飞 {min_to_time_str(earliest_main_dep)}"
            )}
        main_dep_min = forced_main_dep
    else:
        main_dep_min = earliest_main_dep
    main_arr_min = main_dep_min + plan_min

    # ---------- 值勤 ----------
    if day_flights:
        first_dep_abs = (
            (day_flights[0]["dep_date"] - target_date).days * 1440
            + day_flights[0]["dep_min"]
        )
        duty_start = first_dep_abs - 120
    elif ferry_info is not None:
        duty_start = ferry_info["dep_min"] - 120
    else:
        duty_start = main_dep_min - 120
    duty_end = main_arr_min + 60
    duty_total = duty_end - duty_start

    old_flight = 0
    for f in day_flights:
        d = f["dep_min"]
        a = f["arr_min"]
        if a < d:
            a += 1440
        old_flight += (a - d)

    ferry_flight = ferry_info["ferry_min"] if ferry_info else 0

    # ---------- 前日休息 ----------
    prev_day = target_date - timedelta(days=1)
    rest_ok = True
    rest_note = ""
    prev_day_flights = [f for f in reg_flights if f["dep_date"] == prev_day]
    if prev_day_flights:
        prev_last = prev_day_flights[-1]
        prev_duty_end_abs = (
            (prev_last["arr_date"] - target_date).days * 1440
            + prev_last["arr_min"] + 60
        )
        rest_hours = (duty_start - prev_duty_end_abs) / 60
        if rest_hours < 10:
            rest_ok = False
            rest_note = (
                f"前一日值勤结束 {prev_last['arr_date'].strftime('%m月%d日')} "
                f"{min_to_time_str(prev_duty_end_abs)}，"
                f"当日值勤开始 {min_to_time_str(duty_start)}，"
                f"休息仅 {rest_hours:.1f}h < 10h"
            )

    # ---------- 组装航段 ----------
    segments = []
    skipped_ids = set()
    if merged_info:
        for f in merged_info["skipped"]:
            skipped_ids.add((f["dep_date"], f["dep_min"], f["dep_icao"], f["arr_icao"]))

    for f in reg_flights:
        key = (f["dep_date"], f["dep_min"], f["dep_icao"], f["arr_icao"])
        if key in skipped_ids:
            continue
        tag = ""
        if is_ferry_use(f.get("use", "")):
            tag = "原调机"
        segments.append({
            "date": f["dep_date"], "dep_city": f["dep_city"], "dep_icao": f["dep_icao"],
            "dep_min": f["dep_min"], "arr_min": f["arr_min"],
            "arr_city": f["arr_city"], "arr_icao": f["arr_icao"],
            "tag": tag, "use": f.get("use", ""),
        })

    if ferry_info and ferry_info["ferry_min"] > 0:
        fd_abs = ferry_info["dep_min"]
        fa_abs = ferry_info["arr_min"]
        tag = "调机(合并)" if merged_info else "新增调机"
        segments.append({
            "date": target_date + timedelta(days=fd_abs // 1440),
            "dep_city": ferry_info["from_city"], "dep_icao": ferry_info["from_icao"],
            "dep_min": fd_abs % 1440, "arr_min": fa_abs % 1440,
            "arr_city": dep_city, "arr_icao": dep_icao_req,
            "tag": tag, "use": "调机",
        })

    segments.append({
        "date": target_date + timedelta(days=main_dep_min // 1440),
        "dep_city": dep_city, "dep_icao": dep_icao_req,
        "dep_min": main_dep_min % 1440, "arr_min": main_arr_min % 1440,
        "arr_city": arr_city, "arr_icao": arr_icao_req,
        "tag": "新增", "use": "",
    })

    # ---------- 回程调机 ----------
    # 找新需求完成后第一个原计划段，若出发地≠新增到达地，自动加回程调机
    after_main = []
    for s in segments:
        seg_abs_dep = (s["date"] - target_date).days * 1440 + s["dep_min"]
        if seg_abs_dep <= main_arr_min:
            continue
        if s["tag"] != "":
            continue
        after_main.append((seg_abs_dep, s))

    after_main.sort(key=lambda x: x[0])
    return_ferry_info = None
    if after_main:
        next_seg = after_main[0][1]
        next_dep_icao = next_seg["dep_icao"]
        if next_dep_icao != arr_icao_req:
            ret_min = estimate_flight_minutes(arr_icao_req, next_dep_icao)
            if ret_min is not None:
                ret_plan_min = plan_minutes(ret_min)
                ret_dep_abs = main_arr_min + transit_minutes(arr_icao_req)
                ret_arr_abs = ret_dep_abs + ret_plan_min
                segments.append({
                    "date": target_date + timedelta(days=ret_dep_abs // 1440),
                    "dep_city": arr_city, "dep_icao": arr_icao_req,
                    "dep_min": ret_dep_abs % 1440, "arr_min": ret_arr_abs % 1440,
                    "arr_city": next_seg["dep_city"], "arr_icao": next_dep_icao,
                    "tag": "调机(回程)", "use": "调机",
                })
                return_ferry_info = {
                    "from_city": arr_city, "from_icao": arr_icao_req,
                    "to_city": next_seg["dep_city"], "to_icao": next_dep_icao,
                    "ferry_min": ret_min, "ferry_plan_min": ret_plan_min,
                    "dep_min": ret_dep_abs, "arr_min": ret_arr_abs,
                }
                ferry_flight += ret_min

    total_flight = old_flight + ferry_flight + flight_min

    segments.sort(key=lambda x: (x["date"], x["dep_min"]))

    # ★★★ 已删除："新需求之后的调机优化" 循环 ★★★
    # 原来这里有个 for 循环，会把新需求后第一个调机段改成从新增到达地出发
    # 那是错的——因为飞机可能刚刚从某地飞到该调机段的出发地（如 B652S 4 号）
    optimizations = []  # 保留空列表以兼容 UI

    min_date = target_date - timedelta(days=1)
    max_date = target_date + timedelta(days=1)
    visible_segments = [s for s in segments if min_date <= s["date"] <= max_date]

    return {
        "reg": reg, "segments": visible_segments, "day_flights": day_flights,
        "duty_start": duty_start, "duty_end": duty_end, "duty_total": duty_total,
        "old_flight": old_flight, "ferry_info": ferry_info,
        "return_ferry_info": return_ferry_info,
        "merged_info": merged_info,
        "transit": transit, "main_dep": main_dep_min, "main_arr": main_arr_min,
        "new_duty_end": duty_end, "new_duty_total": duty_total,
        "new_flight_total": total_flight,
        "ok_duty": duty_total <= DUTY_MAX_MIN,
        "ok_flight": total_flight <= FLIGHT_MAX_MIN,
        "rest_ok": rest_ok, "rest_note": rest_note,
        "plan_min": plan_min, "flight_min": flight_min,
        "optimizations": optimizations,
    }

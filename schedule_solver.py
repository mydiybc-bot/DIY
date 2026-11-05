"""門市排班工作台：程式排初稿（2026-11-05 經營者「程式先排、AI 只做它擅長的事」）

用 Google OR-Tools CP-SAT 在幾秒內排出整月班表：硬規則（每天人數、指休、兼職可上日、一天一班、七休一）
一定守住；正職時數、例假／休息天數、連上上限、晚接早、兼職最低時數盡量達標，做不到的列成缺口。
AI 只負責把店長寫的備註、店規翻成這裡看得懂的規則（rules），不參與排班本身。
"""
from __future__ import annotations

import time

WEEKEND = ("六", "日")

# 罰分（越大越不能違反；同一個解裡先滿足大的）
# 順序照排班規則：勞基法與休假天數 ＞ 本店連上上限 ＞ 每天人數（補不滿就列缺口，不硬塞違規人力）＞ 其他
W_REST7 = 300000     # 連上超過 6 天（七休一）
W_WEEK = 200000      # 正職完整一週少於 1 例 1 休（每缺 1 天）
W_OFF = 150000       # 正職例假＋休息天數跟本月基準不一樣（每天；不可少排）
W_HOL = 150000       # 正職國定假日天數不足（每天）
W_CONSEC = 120000    # 超過本店連上上限（每多一天的視窗）
W_SHORT = 100000     # 某天在店人數比設定少（每缺 1 人；檢查程式算人頭）
W_SHIFT = 20000      # 某班沒人（人頭夠但班別對不上；檢查只提醒）
W_FORCE = 8000       # 備註指定要上班卻沒排（每天）
W_APART = 60000      # 備註說不能同一天上班卻排在一起（每天）：高於正職時數、低於缺人（2026-11-05 模擬：原本 8000 輸給正職時數，店規 11 天沒守）
W_REST11 = 6000      # 下班到隔天上班不到 11 小時（每次）
W_LATE = 6000        # 晚接早超過本店上限（每次）
W_PT_MIN = 60        # 兼職低於最低時數（每分鐘）
W_FT_HRS = 50        # 正職時數跟目標差（每分鐘）
W_FT_LEND = 2        # 兼職優先方案：正職時數不足（每分鐘；挪到其他月份補，比多排人便宜）
W_EXTRA = 5000       # 某班多排人（每多 1 人；浪費人力）：比兼職目標時數優先，只輸給正職標準時數、兼職最低時數（2026-11-05 經營者：設定 1 人就排 1 人）
W_PT_TGT = 4         # 兼職跟目標時數差（每分鐘）
W_PREFER = 300       # 備註偏好的班別沒排到（每次）
W_HOL_DATE = 1500    # 國定假日沒放在政府公告的那天（每天）
W_SWITCH = 30        # 前後兩天早晚班交換（讓班表穩定一點）


def _mins(t):
    t = str(t)
    nxt = t.startswith("次日")
    t = t.replace("次日", "")
    h, m = t.split(":")
    return int(h) * 60 + int(m) + (1440 if nxt else 0)


def shift_info(cfg):
    """班別 → {name: {segs, minutes, start, end, kind}}；kind＝早／中／晚／''"""
    out = {}
    for r in cfg.get("shifts") or []:
        segs = [list(x) for x in r.get("segs") or [] if x and x[0] and x[1]]
        if not r.get("name") or not segs:
            continue
        mins = sum(_mins(b) - _mins(a) for a, b in segs)
        kind = next((k for k in "早中晚" if k in r["name"]), "")
        out[r["name"]] = {"segs": segs, "minutes": mins, "start": _mins(segs[0][0]), "end": _mins(segs[-1][1]),
                          "kind": kind, "wd": int(r.get("wd") or 0), "we": int(r.get("we") or 0)}
    return out


def demand_of(info, shifts, rules, hol=()):
    """每天每班要幾人 {日: {班別: 人數}}（平日照平日、六日與特殊日「當假日」照假日；備註可改某天某班人數）"""
    need = {}
    for d in info["days"]:
        wk = d["wd"] in WEEKEND or d["d"] in hol
        need[d["d"]] = {n: (s["we"] if wk else s["wd"]) for n, s in shifts.items() if (s["we"] if wk else s["wd"])}
    for r in rules:
        if r["type"] in ("demand", "extra") and r.get("shift") in shifts:
            for dd in r["days"]:
                cur = need[dd].get(r["shift"], 0)
                need[dd][r["shift"]] = max(0, int(r["count"]) if r["type"] == "demand" else cur + int(r["count"]))
    return need


def solve(info, cfg, ms, rules=None, hol_days=(), prev_tail=None, fixed_minutes=None, time_limit=20.0, mode="std"):
    """
    info：104 範本資訊（days、emps）；cfg：店別基本設定；ms：本月設定（mset_view）
    rules：AI 從備註翻出來的規則（見 RULE_TYPES）；hol_days：這個月政府公告放假的平日日期
    prev_tail：{工號: 上個月最後連續上班幾天}；fixed_minutes：{(工號, 日): 固定格子的分鐘數}（rules 的 fixed_cell）
    mode（2026-11-06 經營者：排兩種給店長比較）：
      "ft"＝正職優先：正職排滿標準時數，兼職只補剩下的班（連最低時數都不必滿足）
      "pt"＝兼職優先：兼職先排到目標時數（沒填目標就用最低時數），正職不夠的時數挪到其他月份（多出的休假標「彈性挪假」）
      "std"＝原本的：正職排滿＋兼職排到最低時數
    回傳 {"schedule": {工號: {日: 格子}}, "gaps": [...], "nuo": {工號: [日]}, "stats": {...}}
    """
    from ortools.sat.python import cp_model

    t0 = time.time()
    rules = rules or []
    prev_tail = prev_tail or {}
    fixed_minutes = fixed_minutes or {}
    shifts = shift_info(cfg)
    days = [d["d"] for d in info["days"]]
    wd_of = {d["d"]: d["wd"] for d in info["days"]}
    nd = len(days)
    need = demand_of(info, shifts, rules, {int(x["day"]) for x in ms.get("special") or [] if x.get("as_holiday")})
    caps = {}
    for r in rules:
        if r["type"] == "demand" and r.get("shift") in shifts:
            for dd in r["days"]:
                caps.setdefault(dd, {})[r["shift"]] = int(r["count"])
    base = ms.get("base") or {}
    std_min = int(round(float(base.get("std_hours") or 0) * 60))
    off_days = base.get("off_days")
    holidays = int(base.get("holidays") or 0)
    marks = ms.get("marks") or {}
    avail = ms.get("avail_shift") or {}
    max_consec = int(cfg.get("max_consec") or 5)
    late_limit = cfg.get("late_early_limit")
    late_limit = int(late_limit) if late_limit not in (None, "") else None

    emps = []
    for e in info["emps"]:
        v = (ms.get("emps") or {}).get(e["id"]) or {}
        typ = v.get("type") or "正職"
        tgt = v.get("target") or (v.get("target_hint") if typ == "正職" else None)
        tgt = float(tgt) if tgt not in (None, "") else (std_min / 60 if typ == "正職" and std_min else None)
        mn = v.get("min_hours")
        emps.append({"id": e["id"], "name": e["name"], "ft": typ == "正職",
                     "target": int(round(tgt * 60)) if tgt else None,
                     "min": int(round(float(mn) * 60)) if mn not in (None, "") else None})

    # 每人每天能上哪些班
    off_rule = {(r["emp"], dd) for r in rules if r["type"] == "off" for dd in r["days"]}
    only = {}
    for r in rules:
        if r["type"] == "only_shift":
            for dd in (r.get("days") or days):
                only[(r["emp"], dd)] = set(r["shifts"]) & set(shifts)
    allowed = {}
    pto, fixed = {}, {}
    for e in emps:
        mk = marks.get(e["id"]) or {}
        can = {int(d) for d, v in mk.items() if v == "可上"}
        for d in days:
            v = mk.get(str(d), "")
            if (e["id"], d) in fixed_minutes:
                fixed[(e["id"], d)] = fixed_minutes[(e["id"], d)]
                allowed[(e["id"], d)] = []
                continue
            if v == "特休":
                pto[(e["id"], d)] = True
                allowed[(e["id"], d)] = []
                continue
            if v == "指休" or (e["id"], d) in off_rule:
                allowed[(e["id"], d)] = []
                continue
            ok = [n for n in shifts if need[d].get(n)]
            if not e["ft"] and can:
                if d not in can:
                    ok = []
                elif avail.get(e["id"], {}).get(str(d)) in shifts:
                    ok = [n for n in ok if n == avail[e["id"]][str(d)]]
            lim = only.get((e["id"], d))
            if lim and [n for n in ok if n in lim]:
                ok = [n for n in ok if n in lim]
            allowed[(e["id"], d)] = ok

    def pto_minutes(d):
        """特休那天照常排班：取這天時數最接近 8 小時的班（早班優先）"""
        c = [n for n in shifts if need[d].get(n)] or list(shifts)
        return min(c, key=lambda n: (abs(shifts[n]["minutes"] - 480), "早中晚".find(shifts[n]["kind"]) % 3))

    m = cp_model.CpModel()
    x = {}
    for e in emps:
        for d in days:
            for n in allowed[(e["id"], d)]:
                x[(e["id"], d, n)] = m.NewBoolVar(f"x_{e['id']}_{d}_{n}")
    pen = []

    def work(eid, d):
        """這天有沒有上班（含特休、固定格子）：回傳 (變數加總, 常數)"""
        if (eid, d) in pto or (eid, d) in fixed:
            return [], 1
        return [x[(eid, d, n)] for n in allowed[(eid, d)]], 0

    for e in emps:
        for d in days:
            v, _ = work(e["id"], d)
            if len(v) > 1:
                m.Add(sum(v) <= 1)

    # 每天人數：總人頭不可少、也不要多（多＝浪費人力）；各班別盡量對上
    for d in days:
        tot_need = sum(need[d].values())
        if not tot_need:
            continue
        allv = [x[(e["id"], d, n)] for e in emps for n in allowed[(e["id"], d)]]
        short = m.NewIntVar(0, tot_need, f"short_{d}")
        extra = m.NewIntVar(0, len(emps), f"extra_{d}")
        m.Add(sum(allv) + short >= tot_need)
        m.Add(sum(allv) - tot_need <= extra)
        extra2 = m.NewIntVar(0, len(emps), f"extra2_{d}")    # 同一天多 2 人以上再加扣，讓多出來的人次分散
        m.Add(sum(allv) - tot_need - 1 <= extra2)
        pen += [(short, W_SHORT), (extra, W_EXTRA), (extra2, W_EXTRA)]
        for n, k in need[d].items():
            v = [x[(e["id"], d, n)] for e in emps if (e["id"], d, n) in x]
            miss = m.NewIntVar(0, k, f"miss_{d}_{n}")
            m.Add(sum(v) + miss >= k)
            pen.append((miss, W_SHIFT))
        # 店長明講「某幾天某班只排 N 人」（demand 規則）＝上限，多的人不可以塞進這班
        # 2026-11-05 模擬：店規「週三晚班只排 1 人」，正職時數要排滿，多出來的人被塞到週三晚班
        for n, k in caps.get(d, {}).items():
            v = [x[(e["id"], d, n)] for e in emps if (e["id"], d, n) in x]
            if v:
                m.Add(sum(v) <= k)

    for e in emps:
        eid = e["id"]
        tail = int(prev_tail.get(eid) or 0)
        seq = [None] * min(tail, 6) + days          # None＝上個月最後幾天（都在上班）

        def wv(i):
            if seq[i] is None:
                return [], 1
            return work(eid, seq[i])

        # 七休一＋本店連上上限
        for L, wgt, lim in ((7, W_REST7, 6), (max_consec + 1, W_CONSEC, max_consec)):
            if L > 7 and lim >= 6:
                continue
            for i in range(0, len(seq) - L + 1):
                vs, cs = [], 0
                for j in range(i, i + L):
                    a, b = wv(j)
                    vs += a
                    cs += b
                if not vs and cs <= lim:
                    continue
                over = m.NewIntVar(0, L, f"cons{L}_{eid}_{i}")
                m.Add(sum(vs) + cs - lim <= over)
                pen.append((over, wgt))

        # 下班到隔天上班不到 11 小時、晚接早
        late_pairs = []
        for i in range(nd - 1):
            d1, d2 = days[i], days[i + 1]
            for n1 in allowed[(eid, d1)]:
                for n2 in allowed[(eid, d2)]:
                    gap = shifts[n2]["start"] + 1440 - shifts[n1]["end"]
                    le = shifts[n1]["kind"] == "晚" and shifts[n2]["kind"] == "早"
                    if gap >= 11 * 60 and not le:
                        continue
                    b = m.NewBoolVar(f"pair_{eid}_{d1}_{n1}_{n2}")
                    m.Add(x[(eid, d1, n1)] + x[(eid, d2, n2)] - 1 <= b)
                    if gap < 11 * 60:
                        pen.append((b, W_REST11))
                    if le:
                        late_pairs.append(b)
        if late_pairs and late_limit is not None:
            over = m.NewIntVar(0, nd, f"late_{eid}")
            m.Add(sum(late_pairs) - late_limit <= over)
            pen.append((over, W_LATE))

        # 時數
        terms = [x[(eid, d, n)] * shifts[n]["minutes"] for d in days for n in allowed[(eid, d)]]
        const = sum(shifts[pto_minutes(d)]["minutes"] for d in days if (eid, d) in pto) + \
            sum(fixed.get((eid, d), 0) for d in days)
        H = m.NewIntVar(0, 31 * 24 * 60, f"H_{eid}")
        m.Add(H == sum(terms) + const)
        # 時數目標：低於／高於目標各自的罰分（依 mode）
        if e["ft"]:
            w_under, w_over = (W_FT_LEND, W_FT_HRS) if mode == "pt" else (W_FT_HRS, W_FT_HRS)
            goal = e["target"]
        elif mode == "ft":
            w_under, w_over, goal = 0, 0, None
        elif mode == "pt":
            w_under, w_over, goal = W_PT_MIN, W_PT_TGT, e["target"] or e["min"]
        else:
            w_under, w_over, goal = W_PT_TGT, W_PT_TGT, e["target"]
        if goal:
            under = m.NewIntVar(0, 31 * 24 * 60, f"under_{eid}")
            over = m.NewIntVar(0, 31 * 24 * 60, f"over_{eid}")
            m.Add(H + under >= goal)
            m.Add(H - over <= goal)
            pen += [(under, w_under), (over, w_over)]
        if not e["ft"] and e["min"] and mode == "std":
            short_min = m.NewIntVar(0, e["min"], f"pmin_{eid}")
            m.Add(H + short_min >= e["min"])
            pen.append((short_min, W_PT_MIN))

        # 正職休假：國定假日標記、例假＋休息天數、完整一週至少 1 例 1 休
        if e["ft"]:
            offv, hl = {}, {}
            for d in days:
                v, c = work(eid, d)
                if c:
                    offv[d] = 0
                    continue
                o = m.NewBoolVar(f"off_{eid}_{d}")
                m.Add(o + sum(v) == 1)
                offv[d] = o
                if holidays:
                    h = m.NewBoolVar(f"hol_{eid}_{d}")
                    m.Add(h <= o)
                    hl[d] = h
                    if d not in hol_days:
                        pen.append((h, W_HOL_DATE))
            if holidays:
                hs = m.NewIntVar(0, holidays, f"hs_{eid}")
                m.Add(hs == sum(hl.values()) if hl else hs == 0)
                miss = m.NewIntVar(0, holidays, f"hmiss_{eid}")
                m.Add(hs + miss == holidays)
                pen.append((miss, W_HOL))
            nonhol = [offv[d] for d in days if not isinstance(offv[d], int)]
            if off_days not in (None, ""):
                tot = m.NewIntVar(0, nd, f"tot_{eid}")
                m.Add(tot == sum(nonhol) - sum(hl.values()))
                # 例休不可少排；多休（目標時數比標準少的正職）由時數決定
                dv = m.NewIntVar(0, nd, f"offdev_{eid}")
                m.Add(int(off_days) - tot <= dv)
                pen.append((dv, W_OFF))
            week = []
            for d in days:
                if wd_of[d] == "日":
                    week = [d]
                elif week:
                    week.append(d)
                if len(week) == 7:
                    vs = [offv[w] for w in week if not isinstance(offv[w], int)]
                    hv = [hl[w] for w in week if w in hl]
                    lack = m.NewIntVar(0, 2, f"wk_{eid}_{week[0]}")
                    m.Add(sum(vs) - sum(hv) + lack >= 2)
                    pen.append((lack, W_WEEK))
                    week = []
            e["_off"], e["_hol"] = offv, hl

        # 班別穩定：前後兩天早／晚交換小扣分
        for i in range(nd - 1):
            for k1, k2 in (("早", "晚"), ("晚", "早")):
                a = [x[(eid, days[i], n)] for n in allowed[(eid, days[i])] if shifts[n]["kind"] == k1]
                b = [x[(eid, days[i + 1], n)] for n in allowed[(eid, days[i + 1])] if shifts[n]["kind"] == k2]
                if a and b:
                    s = m.NewBoolVar(f"sw_{eid}_{i}_{k1}")
                    m.Add(sum(a) + sum(b) - 1 <= s)
                    pen.append((s, W_SWITCH))

    # 備註規則
    for r in rules:
        if r["type"] == "work":
            for dd in r["days"]:
                v, c = work(r["emp"], dd)
                if c:
                    continue
                pool = [x[(r["emp"], dd, n)] for n in allowed[(r["emp"], dd)] if not r.get("shift") or n == r["shift"]]
                miss = m.NewBoolVar(f"force_{r['emp']}_{dd}")
                m.Add(sum(pool) + miss >= 1) if pool else m.Add(miss == 1)
                pen.append((miss, W_FORCE))
        elif r["type"] == "prefer_shift":
            for dd in (r.get("days") or days):
                for n in allowed.get((r["emp"], dd), []):
                    if n not in r["shifts"]:
                        pen.append((x[(r["emp"], dd, n)], W_PREFER))
        elif r["type"] == "apart":
            a, b = r["emps"][:2]
            for dd in (r.get("days") or days):
                va, ca = work(a, dd)
                vb, cb = work(b, dd)
                if (va or ca) and (vb or cb):
                    both = m.NewBoolVar(f"apart_{a}_{b}_{dd}")
                    m.Add(sum(va) + ca + sum(vb) + cb - 1 <= both)
                    pen.append((both, W_APART))

    m.Minimize(sum(v * w for v, w in pen))
    sv = cp_model.CpSolver()
    sv.parameters.max_time_in_seconds = time_limit
    sv.parameters.num_search_workers = 2   # 線上機器只有半顆 CPU、512MB，開多也沒用
    st = sv.Solve(m)
    if st not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("程式排不出班表（請檢查班別與人數設定）")

    # 組格子
    sched, gaps = {}, []
    for e in emps:
        eid, row = e["id"], {}
        mk = marks.get(eid) or {}
        offs = e.get("_off") or {}
        hol = e.get("_hol") or {}
        week_has_rest = set()
        for d in days:
            if wd_of[d] == "日":
                week_has_rest = set()
            tag = "[指休]" if mk.get(str(d)) == "指休" else ""
            if (eid, d) in pto:
                row[str(d)] = "[特休]" + pto_minutes(d)
                continue
            if (eid, d) in fixed:
                row[str(d)] = None      # 呼叫端填固定格子
                continue
            n = next((n for n in allowed[(eid, d)] if sv.Value(x[(eid, d, n)])), None)
            if n:
                row[str(d)] = tag + n
            elif e["ft"]:
                if d in hol and sv.Value(hol[d]):
                    row[str(d)] = tag + "國定假日"
                elif "例" not in week_has_rest:
                    row[str(d)] = tag + "例假日"
                    week_has_rest.add("例")
                else:
                    row[str(d)] = tag + "休息日"
            else:
                row[str(d)] = ""        # 兼職沒班整格留空（不標例假／休息）
        sched[eid] = row
    for d in days:
        for n, k in need[d].items():
            got = sum(1 for e in emps if (e["id"], d, n) in x and sv.Value(x[(e["id"], d, n)]))
            if got < k:
                gaps.append({"day": d, "shift": n, "need": k, "got": got})
    hours = {}
    for e in emps:
        tot = sum(shifts[n]["minutes"] for d in days for n in allowed[(e["id"], d)] if sv.Value(x[(e["id"], d, n)]))
        tot += sum(shifts[pto_minutes(d)]["minutes"] for d in days if (e["id"], d) in pto)
        tot += sum(fixed.get((e["id"], d), 0) for d in days)
        hours[e["id"]] = {"name": e["name"], "ft": e["ft"], "hours": tot / 60,
                          "target": e["target"] / 60 if e["target"] else None, "min": e["min"] / 60 if e["min"] else None}
    # 兼職優先：正職少排的天數改標「彈性挪假」（多放的假，之後月份補上班；只標在休息日，例假日不能挪）
    nuo = {}
    if mode == "pt":
        for e in emps:
            if not e["ft"] or not e["target"]:
                continue
            lack = e["target"] - int(round(hours[e["id"]]["hours"] * 60))
            n = int(round(lack / 480)) if lack > 0 else 0
            for d in reversed(days):
                if n <= 0:
                    break
                if sched[e["id"]].get(str(d)) == "休息日":
                    sched[e["id"]][str(d)] = "[彈性挪假]休息日"
                    nuo.setdefault(e["id"], []).append(d)
                    n -= 1
    extras = []
    for d in days:
        k = sum(need[d].values())
        got = sum(1 for e in emps for n in allowed[(e["id"], d)] if sv.Value(x[(e["id"], d, n)]))
        if k and got > k:
            extras.append({"day": d, "need": k, "got": got,
                           "shifts": [n for n in shifts if sum(1 for e in emps if (e["id"], d, n) in x and sv.Value(x[(e["id"], d, n)])) > need[d].get(n, 0)]})
    # 備註規則沒做到的（人不夠時才會破例）→ 報告要寫出來，不能只寫「已套用」
    unmet = []
    worked = lambda eid, d: bool(work(eid, d)[1]) or any(sv.Value(v) for v in work(eid, d)[0])
    for r in rules:
        if r["type"] == "apart":
            a, b = r["emps"][:2]
            both = [d for d in (r.get("days") or days) if worked(a, d) and worked(b, d)]
            if both:
                unmet.append({"rule": r, "days": both})
        elif r["type"] == "work":
            miss = [d for d in r["days"] if not worked(r["emp"], d)]
            if miss:
                unmet.append({"rule": r, "days": miss})
    return {"schedule": sched, "gaps": gaps, "extras": extras, "nuo": nuo, "mode": mode, "hours": hours, "need": need, "unmet": unmet,
            "stats": {"seconds": round(time.time() - t0, 1), "status": sv.StatusName(st), "objective": sv.ObjectiveValue()}}


RULE_TYPES = {
    "off": "某人某幾天不能上班",
    "work": "某人某幾天一定要上班（可指定班別）",
    "only_shift": "某人（某幾天）只能上這些班",
    "prefer_shift": "某人（某幾天）最好上這些班（盡量）",
    "apart": "兩個人（某幾天）不要同一天上班",
    "demand": "某幾天某班要幾個人（取代原本人數）",
    "extra": "某幾天某班多加幾個人",
    "fixed_cell": "某人某天固定用這個格子（上課、會議、跨店支援等自訂時段）",
}

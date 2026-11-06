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
W_LEND = 1500        # 區域一起排：調去同城別家店（每人每天；比多排、缺人便宜，比沒事調店貴）


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
    單店排班（工作台「排初稿」用）。參數意義見 solve_multi 的 unit；回傳那一家店的結果。
    mode（2026-11-06 經營者：排兩種給店長比較）：
      "ft"＝正職優先：正職排滿標準時數，兼職只補剩下的班（連最低時數都不必滿足）
      "pt"＝兼職優先：兼職先排到目標時數（沒填目標就用最低時數），正職不夠的時數挪到其他月份（多出的休假標「彈性挪假」）
      "std"＝原本的：正職排滿＋兼職排到最低時數
    回傳 {"schedule": {工號: {日: 格子}}, "gaps": [...], "nuo": {工號: [日]}, "stats": {...}}
    """
    unit = {"store": 0, "info": info, "cfg": cfg, "ms": ms, "rules": rules, "hol_days": hol_days,
            "prev_tail": prev_tail, "fixed_minutes": fixed_minutes}
    r = solve_multi([unit], mode=mode, lend=False, time_limit=time_limit)
    out = r["stores"][0]
    out["stats"] = r["stats"]
    return out


def solve_multi(units, mode="std", lend=False, time_limit=20.0, workers=2):
    """
    多家店一起排（2026-11-06 經營者「整區一起排」：同城的店一起排，自己店優先，缺人／多人才調店）。
    units：每家店一個 dict：store（店號）、pos（104 崗位名，例「南京店」；調店時寫進格子）、info（104 範本資訊）、
           cfg（店別基本設定）、ms（本月設定 mset_view）、rules（AI 從備註翻的規則）、hol_days（政府公告放假的平日）、
           prev_tail（{工號: 上個月最後連續上班幾天}）、fixed_minutes（{(工號, 日): 固定格子分鐘數}）
    lend：True＝夥伴可以調去同一批的其他店（每調一次扣 W_LEND，所以只在能減少缺人／多排時才調）
    回傳 {"stores": [每家店的結果，順序同 units], "moves": [{store, emp, name, day, to, shift}], "stats": {...}}
    """
    from ortools.sat.python import cp_model

    t0 = time.time()
    U = []
    for u in units:
        info, cfg, ms = u["info"], u["cfg"], u["ms"]
        rules = u.get("rules") or []
        shifts = shift_info(cfg)
        need = demand_of(info, shifts, rules, {int(x["day"]) for x in ms.get("special") or [] if x.get("as_holiday")})
        caps = {}
        for r in rules:
            if r["type"] == "demand" and r.get("shift") in shifts:
                for dd in r["days"]:
                    caps.setdefault(dd, {})[r["shift"]] = int(r["count"])
        base = ms.get("base") or {}
        late_limit = cfg.get("late_early_limit")
        U.append({"store": u["store"], "pos": u.get("pos") or "", "info": info, "ms": ms, "rules": rules, "shifts": shifts,
                  "need": need, "caps": caps, "hol_days": set(u.get("hol_days") or ()), "prev_tail": u.get("prev_tail") or {},
                  "fixed_minutes": u.get("fixed_minutes") or {},
                  "std_min": int(round(float(base.get("std_hours") or 0) * 60)), "off_days": base.get("off_days"),
                  "holidays": int(base.get("holidays") or 0), "max_consec": int(cfg.get("max_consec") or 5),
                  "late_limit": int(late_limit) if late_limit not in (None, "") else None})
    days = [d["d"] for d in U[0]["info"]["days"]]
    wd_of = {d["d"]: d["wd"] for d in U[0]["info"]["days"]}
    nd = len(days)
    SH = {(u["store"], n): s for u in U for n, s in u["shifts"].items()}   # (店, 班別) → 班別資訊

    # 夥伴
    emps = []
    for u in U:
        ms = u["ms"]
        for e in u["info"]["emps"]:
            v = (ms.get("emps") or {}).get(e["id"]) or {}
            typ = v.get("type") or "正職"
            tgt = v.get("target") or (v.get("target_hint") if typ == "正職" else None)
            tgt = float(tgt) if tgt not in (None, "") else (u["std_min"] / 60 if typ == "正職" and u["std_min"] else None)
            mn = v.get("min_hours")
            emps.append({"k": (u["store"], e["id"]), "id": e["id"], "name": e["name"], "u": u, "ft": typ == "正職",
                         "target": int(round(tgt * 60)) if tgt else None,
                         "min": int(round(float(mn) * 60)) if mn not in (None, "") else None})
    E = {e["k"]: e for e in emps}

    # 每人每天能上哪些（店, 班別）
    allowed, pto, fixed = {}, {}, {}
    for u in U:
        rules = u["rules"]
        u["off_rule"] = {(r["emp"], dd) for r in rules if r["type"] == "off" for dd in r["days"]}
        # 區域排班（魔導師對話框加的條件）：no_lend＝某人（某幾天）不調去別店
        u["no_lend"] = {(r["emp"], dd) for r in rules if r["type"] == "no_lend" for dd in (r.get("days") or days)}
        u["only"] = {}
        for r in rules:
            if r["type"] == "only_shift":
                for dd in (r.get("days") or days):
                    u["only"][(r["emp"], dd)] = set(r["shifts"]) & set(u["shifts"])
    for e in emps:
        u, eid, k = e["u"], e["id"], e["k"]
        mk = (u["ms"].get("marks") or {}).get(eid) or {}
        avail = (u["ms"].get("avail_shift") or {}).get(eid) or {}
        can = {int(d) for d, v in mk.items() if v == "可上"}
        for d in days:
            v = mk.get(str(d), "")
            if (eid, d) in u["fixed_minutes"]:
                fixed[k + (d,)] = u["fixed_minutes"][(eid, d)]
                allowed[k + (d,)] = []
                continue
            if v == "特休":
                pto[k + (d,)] = True
                allowed[k + (d,)] = []
                continue
            if v == "指休" or (eid, d) in u["off_rule"]:
                allowed[k + (d,)] = []
                continue
            if not e["ft"] and can and d not in can:
                allowed[k + (d,)] = []
                continue
            ok = [n for n in u["shifts"] if u["need"][d].get(n)]
            if not e["ft"] and can and avail.get(str(d)) in u["shifts"]:
                ok = [n for n in ok if n == avail[str(d)]]
            lim = u["only"].get((eid, d))
            if lim and [n for n in ok if n in lim]:
                ok = [n for n in ok if n in lim]
            opts = [(u["store"], n) for n in ok]
            if lend and not (not e["ft"] and avail.get(str(d))) and (eid, d) not in u["no_lend"]:
                for o in U:
                    if o["store"] != u["store"]:
                        opts += [(o["store"], n) for n in o["shifts"] if o["need"][d].get(n)]
            allowed[k + (d,)] = opts

    def pto_shift(u, d):
        """特休那天照常排班：取這天時數最接近 8 小時的班（早班優先）"""
        c = [n for n in u["shifts"] if u["need"][d].get(n)] or list(u["shifts"])
        return min(c, key=lambda n: (abs(u["shifts"][n]["minutes"] - 480), "早中晚".find(u["shifts"][n]["kind"]) % 3))

    m = cp_model.CpModel()
    x = {}
    for e in emps:
        for d in days:
            for (s, n) in allowed[e["k"] + (d,)]:
                x[e["k"] + (d, s, n)] = m.NewBoolVar("")
    pen = []

    def work(k, d):
        """這天有沒有上班（含特休、固定格子）：回傳 (變數清單, 常數)"""
        if k + (d,) in pto or k + (d,) in fixed:
            return [], 1
        return [x[k + (d, s, n)] for s, n in allowed[k + (d,)]], 0

    def opts_where(k, d, cond):
        return [x[k + (d, s, n)] for s, n in allowed[k + (d,)] if cond(s, n)]

    for e in emps:
        for d in days:
            v, _ = work(e["k"], d)
            if len(v) > 1:
                m.Add(sum(v) <= 1)
            for s, n in allowed[e["k"] + (d,)]:
                if s != e["u"]["store"]:
                    pen.append((x[e["k"] + (d, s, n)], W_LEND))

    # 每家店每天人數：總人頭不可少、也不要多（多＝浪費人力）；各班別盡量對上；店長指定的上限一定守
    by_sd = {}
    for kk, v in x.items():
        by_sd.setdefault((kk[3], kk[2]), []).append((kk[4], v))
    for u in U:
        for d in days:
            tot_need = sum(u["need"][d].values())
            got = by_sd.get((u["store"], d), [])
            if tot_need:
                allv = [v for _, v in got]
                short = m.NewIntVar(0, tot_need, "")
                extra = m.NewIntVar(0, len(emps), "")
                extra2 = m.NewIntVar(0, len(emps), "")   # 同一天多 2 人以上再加扣，讓多出來的人次分散
                m.Add(sum(allv) + short >= tot_need)
                m.Add(sum(allv) - tot_need <= extra)
                m.Add(sum(allv) - tot_need - 1 <= extra2)
                pen += [(short, W_SHORT), (extra, W_EXTRA), (extra2, W_EXTRA)]
                for n, c in u["need"][d].items():
                    vv = [v for nn, v in got if nn == n]
                    miss = m.NewIntVar(0, c, "")
                    m.Add(sum(vv) + miss >= c)
                    pen.append((miss, W_SHIFT))
            for n, c in u["caps"].get(d, {}).items():
                vv = [v for nn, v in got if nn == n]
                if vv:
                    m.Add(sum(vv) <= c)

    for e in emps:
        k, u = e["k"], e["u"]
        tail = int(u["prev_tail"].get(e["id"]) or 0)
        seq = [None] * min(tail, 6) + days          # None＝上個月最後幾天（都在上班）

        def wv(i):
            if seq[i] is None:
                return [], 1
            return work(k, seq[i])

        # 七休一＋本店連上上限
        for L, wgt, lim in ((7, W_REST7, 6), (u["max_consec"] + 1, W_CONSEC, u["max_consec"])):
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
                over = m.NewIntVar(0, L, "")
                m.Add(sum(vs) + cs - lim <= over)
                pen.append((over, wgt))

        # 下班到隔天上班不到 11 小時、晚接早（同一天只會上一個班，依下班／上班時間分組加總即可）
        late_pairs = []
        for i in range(nd - 1):
            d1, d2 = days[i], days[i + 1]
            ends = {SH[o]["end"] for o in allowed[k + (d1,)]}
            starts = {SH[o]["start"] for o in allowed[k + (d2,)]}
            for te in ends:
                for ts in starts:
                    if ts + 1440 - te < 11 * 60:
                        b = m.NewBoolVar("")
                        m.Add(sum(opts_where(k, d1, lambda s, n: SH[(s, n)]["end"] == te)) +
                              sum(opts_where(k, d2, lambda s, n: SH[(s, n)]["start"] == ts)) - 1 <= b)
                        pen.append((b, W_REST11))
            a = opts_where(k, d1, lambda s, n: SH[(s, n)]["kind"] == "晚")
            c = opts_where(k, d2, lambda s, n: SH[(s, n)]["kind"] == "早")
            if a and c:
                b = m.NewBoolVar("")
                m.Add(sum(a) + sum(c) - 1 <= b)
                late_pairs.append(b)
        if late_pairs and u["late_limit"] is not None:
            over = m.NewIntVar(0, nd, "")
            m.Add(sum(late_pairs) - u["late_limit"] <= over)
            pen.append((over, W_LATE))

        # 時數
        terms = [x[k + (d, s, n)] * SH[(s, n)]["minutes"] for d in days for s, n in allowed[k + (d,)]]
        const = sum(u["shifts"][pto_shift(u, d)]["minutes"] for d in days if k + (d,) in pto) + \
            sum(fixed.get(k + (d,), 0) for d in days)
        H = m.NewIntVar(0, 31 * 24 * 60, "")
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
            under = m.NewIntVar(0, 31 * 24 * 60, "")
            over = m.NewIntVar(0, 31 * 24 * 60, "")
            m.Add(H + under >= goal)
            m.Add(H - over <= goal)
            pen += [(under, w_under), (over, w_over)]
        if not e["ft"] and e["min"] and mode == "std":
            short_min = m.NewIntVar(0, e["min"], "")
            m.Add(H + short_min >= e["min"])
            pen.append((short_min, W_PT_MIN))

        # 正職休假：國定假日標記、例假＋休息天數、完整一週至少 1 例 1 休
        if e["ft"]:
            offv, hl = {}, {}
            for d in days:
                v, c = work(k, d)
                if c:
                    offv[d] = 0
                    continue
                o = m.NewBoolVar("")
                m.Add(o + sum(v) == 1)
                offv[d] = o
                if u["holidays"]:
                    h = m.NewBoolVar("")
                    m.Add(h <= o)
                    hl[d] = h
                    if d not in u["hol_days"]:
                        pen.append((h, W_HOL_DATE))
            if u["holidays"]:
                hs = m.NewIntVar(0, u["holidays"], "")
                m.Add(hs == sum(hl.values()) if hl else hs == 0)
                miss = m.NewIntVar(0, u["holidays"], "")
                m.Add(hs + miss == u["holidays"])
                pen.append((miss, W_HOL))
            nonhol = [offv[d] for d in days if not isinstance(offv[d], int)]
            if u["off_days"] not in (None, ""):
                tot = m.NewIntVar(0, nd, "")
                m.Add(tot == sum(nonhol) - sum(hl.values()))
                # 例休不可少排；多休（目標時數比標準少的正職）由時數決定
                dv = m.NewIntVar(0, nd, "")
                m.Add(int(u["off_days"]) - tot <= dv)
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
                    lack = m.NewIntVar(0, 2, "")
                    m.Add(sum(vs) - sum(hv) + lack >= 2)
                    pen.append((lack, W_WEEK))
                    week = []
            e["_hol"] = hl

        # 班別穩定：前後兩天早／晚交換小扣分
        for i in range(nd - 1):
            for k1, k2 in (("早", "晚"), ("晚", "早")):
                a = opts_where(k, days[i], lambda s, n: SH[(s, n)]["kind"] == k1)
                b = opts_where(k, days[i + 1], lambda s, n: SH[(s, n)]["kind"] == k2)
                if a and b:
                    sw = m.NewBoolVar("")
                    m.Add(sum(a) + sum(b) - 1 <= sw)
                    pen.append((sw, W_SWITCH))

    # 備註規則（工號都是同一家店的人）
    for u in U:
        st_ = u["store"]
        for r in u["rules"]:
            if r["type"] == "work":
                for dd in r["days"]:
                    v, c = work((st_, r["emp"]), dd)
                    if c:
                        continue
                    pool = opts_where((st_, r["emp"]), dd, lambda s, n: s == st_ and (not r.get("shift") or n == r["shift"]))
                    miss = m.NewBoolVar("")
                    m.Add(sum(pool) + miss >= 1) if pool else m.Add(miss == 1)
                    pen.append((miss, W_FORCE))
            elif r["type"] == "lend" and lend:    # 區域排班：某人某幾天一定調去某店（可指定那家店的班別）
                for dd in r["days"]:
                    pool = opts_where((st_, r["emp"]), dd, lambda s, n: s == r["to"] and (not r.get("shift") or n == r["shift"]))
                    miss = m.NewBoolVar("")
                    m.Add(sum(pool) + miss >= 1) if pool else m.Add(miss == 1)
                    pen.append((miss, W_FORCE))
            elif r["type"] == "prefer_shift":
                for dd in (r.get("days") or days):
                    for s, n in allowed.get((st_, r["emp"], dd), []):
                        if s == st_ and n not in r["shifts"]:
                            pen.append((x[(st_, r["emp"], dd, s, n)], W_PREFER))
            elif r["type"] == "apart":
                a, b = r["emps"][:2]
                for dd in (r.get("days") or days):
                    va, ca = work((st_, a), dd)
                    vb, cb = work((st_, b), dd)
                    if (va or ca) and (vb or cb):
                        both = m.NewBoolVar("")
                        m.Add(sum(va) + ca + sum(vb) + cb - 1 <= both)
                        pen.append((both, W_APART))

    m.Minimize(sum(v * w for v, w in pen))
    sv = cp_model.CpSolver()
    sv.parameters.max_time_in_seconds = time_limit
    sv.parameters.num_search_workers = workers   # 線上機器只有半顆 CPU、512MB，開多也沒用
    status = sv.Solve(m)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("程式排不出班表（請檢查班別與人數設定）")

    POS = {u["store"]: u["pos"] for u in U}
    pick = {}
    for kk, v in x.items():
        if sv.Value(v):
            pick[kk[:3]] = (kk[3], kk[4])

    def away_cell(s, n):
        """調去別店那天的格子：每段時間標「支援」＋崗位填那家店（104 規則：寫在夥伴自己店的班表）"""
        return ",".join(f"{a}-{b}@支援#{POS[s]}" for a, b in SH[(s, n)]["segs"])

    out, moves = [], []
    for u in U:
        st_ = u["store"]
        marks = u["ms"].get("marks") or {}
        ue = [e for e in emps if e["u"] is u]
        sched, gaps, hours = {}, [], {}
        for e in ue:
            k, row = e["k"], {}
            mk = marks.get(e["id"]) or {}
            hol = e.get("_hol") or {}
            week_has_rest = set()
            mins_ = 0
            for d in days:
                if wd_of[d] == "日":
                    week_has_rest = set()
                tag = "[指休]" if mk.get(str(d)) == "指休" else ""
                if k + (d,) in pto:
                    n = pto_shift(u, d)
                    row[str(d)] = "[特休]" + n
                    mins_ += u["shifts"][n]["minutes"]
                    continue
                if k + (d,) in fixed:
                    row[str(d)] = None      # 呼叫端填固定格子
                    mins_ += fixed[k + (d,)]
                    continue
                p = pick.get(k + (d,))
                if p:
                    s, n = p
                    mins_ += SH[p]["minutes"]
                    if s == st_:
                        row[str(d)] = tag + n
                    else:
                        row[str(d)] = away_cell(s, n)
                        moves.append({"store": st_, "emp": e["id"], "name": e["name"], "day": d, "to": s, "shift": n})
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
            sched[e["id"]] = row
            hours[e["id"]] = {"name": e["name"], "ft": e["ft"], "hours": mins_ / 60,
                              "target": e["target"] / 60 if e["target"] else None, "min": e["min"] / 60 if e["min"] else None}
        here = {}     # 這家店每天每班實際排到的人（含外店調來的）
        for (s0, eid, d), (s, n) in pick.items():
            if s == st_:
                here.setdefault(d, []).append((n, s0 != st_))
        for d in days:
            for n, c in u["need"][d].items():
                got = sum(1 for nn, _ in here.get(d, []) if nn == n)
                if got < c:
                    gaps.append({"day": d, "shift": n, "need": c, "got": got})
        # 兼職優先：正職少排的天數改標「彈性挪假」（多放的假，之後月份補上班；只標在休息日，例假日不能挪）
        nuo = {}
        if mode == "pt":
            for e in ue:
                if not e["ft"] or not e["target"]:
                    continue
                lack = e["target"] - int(round(hours[e["id"]]["hours"] * 60))
                cnt = int(round(lack / 480)) if lack > 0 else 0
                for d in reversed(days):
                    if cnt <= 0:
                        break
                    if sched[e["id"]].get(str(d)) == "休息日":
                        sched[e["id"]][str(d)] = "[彈性挪假]休息日"
                        nuo.setdefault(e["id"], []).append(d)
                        cnt -= 1
        extras = []
        for d in days:
            c = sum(u["need"][d].values())
            got = len(here.get(d, []))
            if c and got > c:
                cnt_n = {}
                for nn, _ in here.get(d, []):
                    cnt_n[nn] = cnt_n.get(nn, 0) + 1
                extras.append({"day": d, "need": c, "got": got,
                               "shifts": [n for n in u["shifts"] if cnt_n.get(n, 0) > u["need"][d].get(n, 0)]})
        guests = [{"day": d, "shift": n} for d, lst in here.items() for n, g in lst if g]
        # 備註規則沒做到的（人不夠時才會破例）→ 報告要寫出來，不能只寫「已套用」
        unmet = []
        worked = lambda eid, d: bool(work((st_, eid), d)[1]) or any(sv.Value(v) for v in work((st_, eid), d)[0])
        for r in u["rules"]:
            if r["type"] == "apart":
                a, b = r["emps"][:2]
                both = [d for d in (r.get("days") or days) if worked(a, d) and worked(b, d)]
                if both:
                    unmet.append({"rule": r, "days": both})
            elif r["type"] == "work":
                miss = [d for d in r["days"] if not worked(r["emp"], d)]
                if miss:
                    unmet.append({"rule": r, "days": miss})
        out.append({"store": st_, "schedule": sched, "gaps": gaps, "extras": extras, "nuo": nuo, "mode": mode,
                    "hours": hours, "need": u["need"], "unmet": unmet, "guests": guests})
    return {"stores": out, "moves": moves,
            "stats": {"seconds": round(time.time() - t0, 1), "status": sv.StatusName(status), "objective": sv.ObjectiveValue()}}


RULE_TYPES = {
    "off": "某人某幾天不能上班",
    "work": "某人某幾天一定要上班（可指定班別）",
    "only_shift": "某人（某幾天）只能上這些班",
    "prefer_shift": "某人（某幾天）最好上這些班（盡量）",
    "apart": "兩個人（某幾天）不要同一天上班",
    "demand": "某幾天某班要幾個人（取代原本人數）",
    "extra": "某幾天某班多加幾個人",
    "fixed_cell": "某人某天固定用這個格子（上課、會議、跨店支援等自訂時段）",
    "no_lend": "（區域排班）某人（某幾天）不調去別店",
    "lend": "（區域排班）某人某幾天一定調去某店",
}

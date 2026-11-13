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
W_FT_FAIR = 3        # 兼職優先方案：同店正職裡「不足最多的那一位」的不足時數（每分鐘；讓不足平均分攤）
W_EXTRA = 5000       # 某班多排人（每多 1 人；浪費人力）：比兼職目標時數優先，只輸給正職標準時數、兼職最低時數（2026-11-05 經營者：設定 1 人就排 1 人）
W_PT_OVER = 20       # 兼職優先方案：兼職超過希望時數（每分鐘）。要比正職不足（2＋3）重，排滿兼職後剩下的班先給正職（2026-11-11：原本 4，實測兼職被排到希望時數的 1.5 倍、正職反而不足）
W_PT_TGT = 4         # 兼職跟目標時數差（每分鐘）
W_PREFER = 300       # 備註偏好的班別沒排到（每次）
W_HOL_DATE = 1500    # 國定假日沒放在政府公告的那天（每天）
W_SWITCH = 30        # 前後兩天早晚班交換（讓班表穩定一點）
W_OFFRUN = 1200      # 正職連休 3 天以上（每多 1 天；店長指休／特休自己連成的不算）。2026-11-12 經營者：休假太集中要改善。
                     # 低於國定假日放在公告當天（1500）、多排人（3400 起），不會為了散開休假而多排人或挪國定假日
W_LONE = 400         # 正職休假中間只夾 1 天班（休・上・休）：讓休假平均散開，不要碎成一天一天
W_LEND = 1500        # 區域一起排：調去同城別家店（每人每天；比多排、缺人便宜，比沒事調店貴）

# 2026-11-10 經營者：人力先滿足 週六 ＞ 國定假日 ＞ 週日 ＞ 週五 ＞ 四 ＞ 三 ＞ 二 ＞ 一
# 缺人、某班沒人：順位越前面罰越重（人不夠時缺口留給順位後面的日子）；多排人：順位越前面罰越輕（多出來的人力先放前面的日子）。
# 加減的幅度不跨層：缺 1 人會同時算「人數少」＋「某班沒人」，兩項合計最高（週六）維持原本的 120000、順位越後面越低
# （2026-11-11 修：原本用加的，合計 122400～139200 超過本店連上上限 120000，人不夠時會寧可讓人多連上一天）；多排最低 3400 仍高於調店（1500）。
DAY_RANK = {"六": 8, "日": 6, "五": 5, "四": 4, "三": 3, "二": 2, "一": 1}
RANK_HOL = 7


def day_rank(wd, is_hol):
    """這天的人力順位（8 最優先）：週六 8、國定假日（政府公告放假的平日）7、週日 6、週五～週一 5～1"""
    r = DAY_RANK.get(wd, 1)
    return max(r, RANK_HOL) if is_hol else r


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


def solve(info, cfg, ms, rules=None, hol_days=(), prev_tail=None, fixed_minutes=None, time_limit=20.0, mode="std", fixed_cover=None,
          prev_week=None):
    """
    單店排班（工作台「排初稿」用）。參數意義見 solve_multi 的 unit；回傳那一家店的結果。
    mode（2026-11-06 經營者：排兩種給店長比較）：
      "ft"＝正職優先：正職排滿標準時數，兼職只補剩下的班（連最低時數都不必滿足）
      "pt"＝兼職優先：兼職先排到目標時數（沒填目標就用最低時數），正職不夠的時數挪到其他月份（多出的休假標「彈性挪假」）
      "std"＝原本的：正職排滿＋兼職排到最低時數
    回傳 {"schedule": {工號: {日: 格子}}, "gaps": [...], "nuo": {工號: [日]}, "stats": {...}}
    """
    unit = {"store": 0, "info": info, "cfg": cfg, "ms": ms, "rules": rules, "hol_days": hol_days,
            "prev_tail": prev_tail, "fixed_minutes": fixed_minutes, "fixed_cover": fixed_cover, "prev_week": prev_week}
    r = solve_multi([unit], mode=mode, lend=False, time_limit=time_limit)
    out = r["stores"][0]
    out["stats"] = r["stats"]
    return out


def solve_multi(units, mode="std", lend=False, time_limit=20.0, workers=2):
    """
    多家店一起排（2026-11-06 經營者「整區一起排」：同城的店一起排，自己店優先，缺人／多人才調店）。
    units：每家店一個 dict：store（店號）、pos（104 崗位名，例「南京店」；調店時寫進格子）、info（104 範本資訊）、
           cfg（店別基本設定）、ms（本月設定 mset_view）、rules（AI 從備註翻的規則）、hol_days（政府公告放假的平日）、
           prev_tail（{工號: 上個月最後連續上班幾天}）、fixed_minutes（{(工號, 日): 固定格子分鐘數}）、
           prev_week（{工號: {"off": 上個月落在本月第一週的休假天數, "li": 那幾天有沒有例假日}}；2026-11-13 跨月那週要接起來算 1 例 1 休）
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
        # 2026-11-13：備註指定的自訂班（固定格子）如果時間蓋住某個班大部分（例：10 小時班蓋住早班）→ 那個班這天少排 1 人，
        # 不然會多排一個人。回報用的 need_orig 不變（檢查是算人頭，固定格子的人本來就算在店）。
        need_orig = {d: dict(v) for d, v in need.items()}
        for (_, dd), cand in (u.get("fixed_cover") or {}).items():
            hit = next((n for n in cand if need.get(dd, {}).get(n)), None)
            if hit:
                need[dd][hit] -= 1
                if not need[dd][hit]:
                    del need[dd][hit]
        caps = {}
        for r in rules:
            if r["type"] == "demand" and r.get("shift") in shifts:
                for dd in r["days"]:
                    caps.setdefault(dd, {})[r["shift"]] = int(r["count"])
        base = ms.get("base") or {}
        late_limit = cfg.get("late_early_limit")
        U.append({"store": u["store"], "pos": u.get("pos") or "", "info": info, "ms": ms, "rules": rules, "shifts": shifts,
                  "need": need, "need_orig": need_orig, "caps": caps, "positions": set(info.get("positions") or ()), "hol_days": set(u.get("hol_days") or ()), "prev_tail": u.get("prev_tail") or {},
                  "fixed_minutes": u.get("fixed_minutes") or {},
                  "std_min": int(round(float(base.get("std_hours") or 0) * 60)), "off_days": base.get("off_days"),
                  "holidays": int(base.get("holidays") or 0), "max_consec": int(cfg.get("max_consec") or 5),
                  "prev_week": u.get("prev_week") or {}, "tags": set(info.get("tags") or ()),
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
                         "move": True,   # 2026-11-11 經營者：全部夥伴都可調店（後台固定）
                         "dflt": bool(v.get("_default")),
                         "target": int(round(tgt * 60)) if tgt else None,
                         "min": int(round(float(mn) * 60)) if mn not in (None, "") else None})
    E = {e["k"]: e for e in emps}

    # 每人每天能上哪些（店, 班別）
    allowed, pto, fixed = {}, {}, {}
    pos_block = set()
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
    roster = {e["k"] for e in emps}   # (店, 工號)：這人在哪幾家店的 104 名單上
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
                # 2026-11-13：可上的班這天不用人（例：人數填 0 的「平日晚班6h」）→ 改看同一種（早／中／晚）有要人的班，不然這天整個排不進去
                kd = u["shifts"][avail[str(d)]]["kind"]
                ok = [n for n in ok if n == avail[str(d)]] or ([n for n in ok if u["shifts"][n]["kind"] == kd] if kd else [])
            lim = u["only"].get((eid, d))
            if lim and [n for n in ok if n in lim]:
                ok = [n for n in ok if n in lim]
            opts = [(u["store"], n) for n in ok]
            if lend and e["move"] and (eid, d) not in u["no_lend"]:
                # 兼職這天指定了班別（例：只能晚班）→ 別家店也只排同一種（早／中／晚；看不出種類就要同名）
                want = u["shifts"].get(avail.get(str(d))) if (not e["ft"] and can) else None
                for o in U:
                    # 2026-11-10：來源店的 104 範本要有那家店的崗位才能調（沒有的話寫得進班表、匯不進 104）
                    if o["store"] != u["store"] and o["pos"] and u["positions"] and o["pos"] not in u["positions"]:
                        pos_block.add((u["store"], o["store"]))
                        continue
                    # 2026-11-13 經營者「混合做法」：人已在那家店的 104 名單上 → 班寫在那家店（那家店 104 看得到），不在自己店寫支援
                    if o["store"] != u["store"] and (o["store"], eid) not in roster:
                        opts += [(o["store"], n) for n in o["shifts"] if o["need"][d].get(n) and (
                            not want or (o["shifts"][n]["kind"] == want["kind"] if want["kind"] else n == avail.get(str(d))))]
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
    pen, pen_sp = [], []    # pen＝原本的規則；pen_sp＝休假散開（第二段才算）
    ft_under = {}
    # 2026-11-10：同一個工號掛在同區兩家店的名單（跨店支援的夥伴）＝同一個人，同一天只能在一家店上班
    same = {}
    for e in emps:
        same.setdefault(e["k"][1], []).append(e)
    for grp in same.values():
        if len(grp) < 2:
            continue
        for d in days:
            vs, const = [], 0
            for e in grp:
                kd = e["k"] + (d,)
                if kd in pto or kd in fixed:
                    const += 1
                vs += [x[kd + (s_, n_)] for s_, n_ in allowed[kd]]
            if vs:
                m.Add(sum(vs) <= max(0, 1 - const))
    # 同一人掛在兩店名單（2026-11-11 經營者：兩家店的店長會各自設定他的時數，照店長的設定排）：
    # ・店長在本月設定存過這位夥伴的那幾份 → 各照各店設定的目標／最低時數排
    # ・沒存過（只有系統預設值）的那幾份 → 不追時數、不列入正職平均分攤、不標可支援
    # ・每一份都沒存過 → 只認第一份，時數用各店加總去比
    home = {}
    for eid_, grp in same.items():
        if len(grp) > 1:
            setd = [g for g in grp if not g["dflt"]]
            for g in grp:
                g["alias"] = (g["dflt"] if setd else g is not grp[0])
            if not setd:
                home[eid_] = grp[0]
    # 2026-11-13 經營者「混合做法」：同一人掛好幾家店名單時，例假／休息日／國定假日只標在「自己店」那一筆（店長存過、設正職的那筆；
    # 都沒存過就第一筆），而且要看他整天在各店有沒有上班——原本每筆各算各的，別店名單那筆排了班、自己店照樣標休息日，104 匯入會擋
    rest_owner = {}
    for eid_, grp in same.items():
        fts = [g for g in grp if g["ft"]]
        if fts:
            rest_owner[eid_] = next((g for g in fts if not g.get("alias")), fts[0])
    Hk, deferred = {}, []

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
                rk = day_rank(wd_of[d], d in u["hol_days"])
                pen += [(short, W_SHORT - (8 - rk) * 2000), (extra, W_EXTRA - rk * 200), (extra2, W_EXTRA)]
                for n, c in u["need"][d].items():
                    vv = [v for nn, v in got if nn == n]
                    miss = m.NewIntVar(0, c, "")
                    m.Add(sum(vv) + miss >= c)
                    pen.append((miss, W_SHIFT - (8 - rk) * 400))
            for n, c in u["caps"].get(d, {}).items():
                vv = [v for nn, v in got if nn == n]
                if vv:
                    m.Add(sum(vv) <= c)

    for e in emps:
        k, u = e["k"], e["u"]
        grp = same.get(e["id"]) or [e]              # 2026-11-11 經營者：同一人掛兩店名單，連上天數要各店合起來算
        tail = max(int(g["u"]["prev_tail"].get(g["id"]) or 0) for g in grp)
        seq = [None] * min(tail, 6) + days          # None＝上個月最後幾天（都在上班）

        def wv(i):
            if seq[i] is None:
                return [], 1
            vs_, cs_ = [], 0
            for g in grp:
                a_, b_ = work(g["k"], seq[i])
                vs_ += a_
                cs_ += b_
            return vs_, min(cs_, 1)

        # 七休一＋本店連上上限（跨店的人只算一次，上限取各店最嚴的）
        mc = min(g["u"]["max_consec"] for g in grp)
        for L, wgt, lim in (() if e is not grp[0] else ((7, W_REST7, 6), (mc + 1, W_CONSEC, mc))):
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
        Hk[k] = H
        # 時數目標：低於／高於目標各自的罰分（依 mode）
        if e["ft"]:
            w_under, w_over = (W_FT_LEND, W_FT_HRS) if mode == "pt" else (W_FT_HRS, W_FT_HRS)
            goal = e["target"]
        elif mode == "ft":
            w_under, w_over, goal = 0, 0, None
        elif mode == "pt":
            w_under, w_over, goal = W_PT_MIN, W_PT_OVER, e["target"] or e["min"]
        else:
            w_under, w_over, goal = W_PT_TGT, W_PT_TGT, e["target"]
        if e.get("alias"):
            goal = None
        elif e["k"][1] in home:
            deferred.append((e, w_under, w_over, goal))
            goal = None
        if goal:
            under = m.NewIntVar(0, 31 * 24 * 60, "")
            over = m.NewIntVar(0, 31 * 24 * 60, "")
            m.Add(H + under >= goal)
            m.Add(H - over <= goal)
            pen += [(under, w_under), (over, w_over)]
            if e["ft"] and mode == "pt":
                ft_under.setdefault(e["u"]["store"], []).append(under)
        if not e["ft"] and e["min"] and mode == "std" and not e.get("alias"):
            short_min = m.NewIntVar(0, e["min"], "")
            m.Add(H + short_min >= e["min"])
            pen.append((short_min, W_PT_MIN))

        # 正職休假：國定假日標記、例假＋休息天數、完整一週至少 1 例 1 休
        if e["ft"] and rest_owner.get(e["id"], e) is e:
            offv, hl = {}, {}
            for d in days:
                v, c = [], 0
                for g in same.get(e["id"]) or [e]:     # 各店名單合起來：任何一家店有班就不是休假
                    a_, b_ = work(g["k"], d)
                    v += a_
                    c += b_
                c = min(c, 1)
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
            # 2026-11-13 經營者「跨月份要修」：月初不是週日時，第一週（104 以週日為一週第一天）前幾天在上個月，
            # 要接上個月那幾天的休假一起算 1 例 1 休（12 月模擬：11/29～12/5 有 7 人一週上 6 天、48 小時）
            pw = u["prev_week"].get(e["id"])
            first = []
            for d in days:
                if wd_of[d] == "日":
                    break
                first.append(d)
            if first and pw is not None:
                vs = [offv[w] for w in first if not isinstance(offv[w], int)]
                hv = [hl[w] for w in first if w in hl]
                lack = m.NewIntVar(0, 2, "")
                m.Add(sum(vs) - sum(hv) + int(pw.get("off") or 0) + lack >= 2)
                pen.append((lack, W_WEEK))
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
            # 2026-11-12 經營者：休假不要集中 → 連休最多 2 天（第 3 天起每天扣分）、不要「休・上・休」
            # 整段都是店長指休／備註不能上的不扣（那是店長排的）；特休、固定格子當作有上班
            forced = {d for d in days if not isinstance(offv[d], int) and not any(work(g["k"], d)[0] for g in same.get(e["id"]) or [e])}
            for i in range(nd - 2):
                w3 = days[i:i + 3]
                if any(isinstance(offv[d], int) for d in w3) or all(d in forced for d in w3):
                    continue
                b3 = m.NewBoolVar("")
                m.Add(sum(offv[d] for d in w3) - 2 <= b3)
                pen_sp.append((b3, W_OFFRUN))
            for i in range(1, nd - 1):
                p_, d_, n_ = days[i - 1], days[i], days[i + 1]
                if isinstance(offv[d_], int) or d_ in forced or isinstance(offv[p_], int) or isinstance(offv[n_], int):
                    continue
                if p_ in forced and n_ in forced:
                    continue
                b1 = m.NewBoolVar("")
                m.Add(offv[p_] + offv[n_] - offv[d_] - 1 <= b1)
                pen_sp.append((b1, W_LONE))
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

    for e, w_under, w_over, goal in deferred:
        if not goal:
            continue
        tot = sum(Hk[g["k"]] for g in same[e["k"][1]])
        under = m.NewIntVar(0, 62 * 24 * 60, "")
        over = m.NewIntVar(0, 62 * 24 * 60, "")
        m.Add(tot + under >= goal)
        m.Add(tot - over <= goal)
        pen += [(under, w_under), (over, w_over)]
        if e["ft"] and mode == "pt":
            ft_under.setdefault(e["u"]["store"], []).append(under)
    # 2026-11-10：兼職優先時，正職在本店排不滿的時數要平均分給同店正職（原本只算總量，會變成一人 168h、一人 16h）
    for vs in ft_under.values():
        if len(vs) > 1:
            mx = m.NewIntVar(0, 31 * 24 * 60, "")
            for v in vs:
                m.Add(mx >= v)
            pen.append((mx, W_FT_FAIR))

    # 2026-11-12：分兩段排。第一段照原本的規則排（人數、時數、休假天數…）；第二段在「原本的分數不可變差」下把休假散開。
    # 一次排的話線上機器（半顆 CPU）會在時限內顧此失彼：實測休假散開了，但時數等原本目標變差。
    main = sum(v * w for v, w in pen)
    m.Minimize(main)
    sv = cp_model.CpSolver()
    sv.parameters.max_time_in_seconds = time_limit * 0.6 if pen_sp else time_limit
    sv.parameters.num_search_workers = workers   # 線上機器只有半顆 CPU、512MB，開多也沒用
    status = sv.Solve(m)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("程式排不出班表（請檢查班別與人數設定）")
    if pen_sp:
        m.Add(main <= int(round(sv.ObjectiveValue())))
        for v in x.values():
            m.AddHint(v, sv.Value(v))
        m.Minimize(main + sum(v * w for v, w in pen_sp))
        sv2 = cp_model.CpSolver()
        sv2.parameters.max_time_in_seconds = max(5.0, time_limit - (time.time() - t0))
        sv2.parameters.num_search_workers = workers
        st2 = sv2.Solve(m)
        if st2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):   # 第二段沒排出來就用第一段的結果
            sv, status = sv2, st2

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
            week_has_rest = {"例"} if (u["prev_week"].get(e["id"]) or {}).get("li") else set()   # 跨月那週上個月已有例假 → 本月這週的休假標休息日
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
                elif e["ft"] and (rest_owner.get(e["id"], e) is not e or any(
                        pick.get(g["k"] + (d,)) or g["k"] + (d,) in pto or g["k"] + (d,) in fixed
                        for g in same.get(e["id"]) or [] if g is not e)):
                    row[str(d)] = ""        # 休假只標在自己店那筆；那天在別店名單上班 → 自己店留空白（104 一人一天一份班表）
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
                if not e["ft"] or not e["target"] or e.get("alias"):
                    continue
                other = sum(sv.Value(Hk[g["k"]]) for g in same.get(e["id"], []) if g is not e) if e["id"] in home else 0   # 各店都沒設定時：扣掉別店名單排到的時數
                lack = e["target"] - int(round(hours[e["id"]]["hours"] * 60)) - other
                cnt = int(lack / 480 + 0.5) if lack > 0 else 0   # 四捨五入（原本 round() 遇 .5 取偶數：少 12h 標 2 天、少 4h 標 0 天）
                for d in reversed(days):
                    if cnt <= 0:
                        break
                    if sched[e["id"]].get(str(d)) == "休息日":
                        # 2026-11-13：這家店的 104 班表沒啟用「彈性挪假」標籤時不貼（貼了匯入會失敗），只記在 nuo、報告提醒
                        sched[e["id"]][str(d)] = "[彈性挪假]休息日" if (not u["tags"] or "彈性挪假" in u["tags"]) else "休息日"
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
                    "hours": hours, "need": u["need_orig"], "unmet": unmet, "guests": guests})
    return {"pos_block": sorted(pos_block), "stores": out, "moves": moves,
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

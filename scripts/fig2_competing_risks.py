#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fig 2：纽扣电池自然排出的累积发生曲线（竞争风险）/ Competing-risks CIF.

为什么不能用 Kaplan-Meier
------------------------
"自然排出"的观察会被"内镜/手术取出"打断，而取出并非随机删失——医生正是因为
判断它不会自行排出才去取。把取出当作删失会系统性高估自然排出率。
因此用 Aalen-Johansen 估计量，把取出作为竞争事件单独建模。

事件定义
--------
  时间原点  误食时刻（入院时刻 − 主诉解析出的就诊延迟）
  关注事件  自然排出：末次"异物在位"片与首次"异物消失"片之间（取区间中点）
  竞争事件  内镜或手术取出（手麻系统记录的手术开始时刻）
  删失      出院时既未确认排出也未取出

排出时间是区间删失的（两次复查之间），本队列区间宽度中位约 20 h。
默认取区间中点，--endpoint right 改用首次阴性片时刻（上界），作敏感性分析。

用法
----
  python3 scripts/fig2_competing_risks.py
  python3 scripts/fig2_competing_risks.py --endpoint right --xmax 168
  python3 scripts/fig2_competing_risks.py --selftest     # 验证估计量实现
"""
import argparse
import os
import sys

import numpy as np
import pandas as pd

import logging
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# 容器/服务器上常无 Helvetica，回退字体的提示会淹没日志；投稿机器上应装 Helvetica 或 Arial
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cohort as C

OUTDIR = "output"
FIGDIR = "figures"
PASSAGE, REMOVAL, CENSORED = 1, 2, 0
REPORT_TIMES = [12, 24, 48, 72]
N_BOOT = 1000
RNG_SEED = 20260913

plt.rcParams.update({
    "font.family": ["Helvetica Neue", "Helvetica", "Arial", "DejaVu Sans"],
    "font.size": 9, "axes.linewidth": 0.8,
    "xtick.major.width": 0.8, "ytick.major.width": 0.8,
    "savefig.bbox": "tight", "svg.fonttype": "none",
})


# ============================================================
# Aalen-Johansen 累积发生函数
# ============================================================
def aalen_johansen(time, event, cause, grid):
    """竞争风险下 cause 的累积发生函数，估计在 grid 上。

    CIF_k(t) = Σ_{t_j ≤ t} S(t_{j-1}) · d_kj / n_j
    其中 S 为全因无事件生存（KM），n_j 为风险集，d_kj 为 t_j 时 cause k 的事件数。
    """
    time = np.asarray(time, float)
    event = np.asarray(event, int)
    order = np.argsort(time)
    time, event = time[order], event[order]

    uniq = np.unique(time[event != CENSORED])
    surv, cif = 1.0, 0.0
    out, gi = np.zeros(len(grid)), 0
    for tj in uniq:
        n_j = int((time >= tj).sum())
        if n_j == 0:
            break
        d_all = int(((time == tj) & (event != CENSORED)).sum())
        d_k = int(((time == tj) & (event == cause)).sum())
        while gi < len(grid) and grid[gi] < tj:
            out[gi] = cif
            gi += 1
        cif += surv * d_k / n_j          # 用 t_j 之前的 S
        surv *= (1 - d_all / n_j)
        while gi < len(grid) and grid[gi] == tj:
            out[gi] = cif
            gi += 1
    out[gi:] = cif
    return out


def bootstrap_ci(time, event, cause, grid, n_boot=N_BOOT, seed=RNG_SEED):
    """百分位 bootstrap 置信带。样本量小，比解析方差更稳妥且易于说明。"""
    rng = np.random.default_rng(seed)
    n = len(time)
    time, event = np.asarray(time, float), np.asarray(event, int)
    curves = np.empty((n_boot, len(grid)))
    for b in range(n_boot):
        idx = rng.integers(0, n, n)
        curves[b] = aalen_johansen(time[idx], event[idx], cause, grid)
    return np.percentile(curves, 2.5, axis=0), np.percentile(curves, 97.5, axis=0)


def permutation_test(df, cause, grid, n_perm=2000, seed=RNG_SEED):
    """组间 CIF 差异的置换检验。

    Gray 检验在 Python 中无成熟实现（R 的 cmprsk 有），故改用置换检验：
    统计量为两组 CIF 曲线差的绝对面积，组标签随机重排以生成零分布。
    在 Methods 中须如实写明用的是置换检验而非 Gray 检验。
    """
    rng = np.random.default_rng(seed)
    groups = df["stratum"].unique()
    if len(groups) != 2:
        return None
    g0, g1 = groups

    def stat(labels):
        a = df[labels == g0]
        b = df[labels == g1]
        if len(a) < 3 or len(b) < 3:
            return 0.0
        ca = aalen_johansen(a["time_h"], a["event"], cause, grid)
        cb = aalen_johansen(b["time_h"], b["event"], cause, grid)
        return np.trapezoid(np.abs(ca - cb), grid)

    obs = stat(df["stratum"].values)
    labels = df["stratum"].values.copy()
    null = np.array([stat(rng.permutation(labels)) for _ in range(n_perm)])
    return obs, float((null >= obs).mean())


# ============================================================
# 数据构建
# ============================================================
def build_events(endpoint="midpoint"):
    xl = C.load_workbook()
    base = xl.parse("病案首页基本信息")
    adm = xl.parse("儿科入院记录")
    xray = xl.parse("X线报告")
    anes = xl.parse("手麻系统信息")

    core = C.build_core_text(adm)
    verified, is_verified = C.load_verified_cohort()
    battery = verified if is_verified else C.ascertain(core, C.BATTERY_KW)

    base = base.copy()
    base["adm_t"] = pd.to_datetime(base["入院日期"], errors="coerce")
    base["dis_t"] = pd.to_datetime(base["出院日期"], errors="coerce")
    admT = dict(zip(base["科研就诊编号"], base["adm_t"]))
    disT = dict(zip(base["科研就诊编号"], base["dis_t"]))
    cc = dict(zip(adm["科研就诊编号"], adm["主诉"]))

    anes = anes.copy()
    anes["t"] = pd.to_datetime(anes["手术开始时间"], errors="coerce")
    procT = (anes[anes["科研就诊编号"].isin(battery)]
             .sort_values("t").groupby("科研就诊编号")["t"].first().to_dict())

    xray = xray.copy()
    xray["t"] = pd.to_datetime(xray["检查时间"], errors="coerce")
    bx = xray[xray["科研就诊编号"].isin(battery)].sort_values("t")

    rows = []
    for v in sorted(battery):
        adm_t = admT.get(v)
        delay = C.parse_delay_hours(cc.get(v, ""))
        if pd.isna(adm_t):
            continue
        ingest = adm_t - pd.Timedelta(hours=delay) if delay is not None else adm_t

        g = bx[bx["科研就诊编号"] == v]
        states = [(r["t"], C.classify_film_state(
            str(r["检查所见"]) + "。" + str(r["检查结论"]))[0]) for _, r in g.iterrows()]
        first_loc = None
        if len(g):
            ft = str(g.iloc[0]["检查所见"]) + " " + str(g.iloc[0]["检查结论"])
            first_loc = C.classify_location(ft)

        last_present = max([t for t, s in states if s == "present"], default=None)
        absents = [t for t, s in states if s == "absent"
                   and (last_present is None or t > last_present)]
        first_absent = min(absents) if absents else None

        t_proc = procT.get(v)
        t_dis = disT.get(v)

        if first_absent is not None:
            if endpoint == "right" or last_present is None:
                lo = last_present if last_present is not None else ingest
                t_event = first_absent if endpoint == "right" else lo + (first_absent - lo) / 2
            else:
                t_event = last_present + (first_absent - last_present) / 2
            ev, t = PASSAGE, t_event
        elif t_proc is not None and pd.notna(t_proc):
            ev, t = REMOVAL, t_proc
        elif t_dis is not None and pd.notna(t_dis):
            ev, t = CENSORED, t_dis
        else:
            continue

        hours = (t - ingest).total_seconds() / 3600
        if hours <= 0 or not np.isfinite(hours):
            continue

        stratum = ("Stomach or duodenum" if first_loc in ("stomach", "duodenum")
                   else "Distal to the duodenum"
                   if first_loc in ("small_bowel", "colorectal", "passed_or_absent")
                   else None)
        rows.append({"科研就诊编号": v, "time_h": hours, "event": ev,
                     "stratum": stratum, "first_loc": first_loc,
                     "interval_h": ((first_absent - last_present).total_seconds() / 3600
                                    if (first_absent is not None and last_present is not None) else None)})
    return pd.DataFrame(rows), is_verified


# ============================================================
# 绘图
# ============================================================
def plot_figure(df, grid, args):
    strata = ["Stomach or duodenum", "Distal to the duodenum"]
    styles = {strata[0]: dict(color="#1a1a1a", ls="-"),
              strata[1]: dict(color="#A0480A", ls="--")}

    fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.3), sharey=True)
    panels = [(PASSAGE, "A", "Spontaneous passage"),
              (REMOVAL, "B", "Endoscopic or surgical removal")]

    for ax, (cause, tag, title) in zip(axes, panels):
        for s in strata:
            sub = df[df["stratum"] == s]
            if len(sub) < 3:
                continue
            cif = aalen_johansen(sub["time_h"], sub["event"], cause, grid)
            lo, hi = bootstrap_ci(sub["time_h"], sub["event"], cause, grid,
                                  n_boot=args.boot)
            st = styles[s]
            ax.step(grid, cif, where="post", lw=1.6, label=f"{s} (n={len(sub)})", **st)
            ax.fill_between(grid, lo, hi, step="post", alpha=0.13,
                            color=st["color"], lw=0)
        ax.set_xlim(0, args.xmax)
        ax.set_ylim(0, 1.0)
        ax.set_xticks(np.arange(0, args.xmax + 1, 12))
        ax.set_xlabel("Hours since ingestion")
        ax.set_title(title, fontsize=10, fontweight="bold", loc="left", pad=10)
        ax.text(-0.13, 1.14, tag, transform=ax.transAxes,
                fontsize=12, fontweight="bold", va="top")
        ax.grid(axis="y", lw=0.4, alpha=0.3)
        ax.spines[["top", "right"]].set_visible(False)

    axes[0].set_ylabel("Cumulative incidence")
    axes[0].legend(frameon=False, fontsize=8.5, loc="upper left")

    # 风险集人数
    at_risk_t = np.arange(0, args.xmax + 1, 12)
    lines = []
    for s in strata:
        sub = df[df["stratum"] == s]
        lines.append(f"{s}: " + "  ".join(
            f"{int((sub['time_h'] >= t).sum()):>3d}" for t in at_risk_t))
    fig.text(0.005, -0.10,
             "Number at risk (per 12 h, both panels)\n" + "\n".join(lines)
             + "\n\nEstimates beyond 48 h rest on few patients; interpret the right-hand"
               " portion of each curve with caution.",
             fontsize=7.5, family="monospace", va="top")

    os.makedirs(FIGDIR, exist_ok=True)
    for ext in ("svg", "png"):
        fig.savefig(os.path.join(FIGDIR, f"fig2_competing_risks.{ext}"),
                    dpi=300 if ext == "png" else None)
    plt.close(fig)


# ============================================================
# 自检：与手算例核对
# ============================================================
def selftest():
    # t=1 cause1, t=2 cause2, t=3 censored, t=4 cause1, t=5 cause1
    time = [1, 2, 3, 4, 5]
    event = [1, 2, 0, 1, 1]
    grid = np.array([1, 2, 3, 4, 5], float)
    got1 = aalen_johansen(time, event, 1, grid)
    got2 = aalen_johansen(time, event, 2, grid)
    exp1 = [0.2, 0.2, 0.2, 0.5, 0.8]
    exp2 = [0.0, 0.2, 0.2, 0.2, 0.2]
    ok1 = np.allclose(got1, exp1)
    ok2 = np.allclose(got2, exp2)
    print("手算例 CIF(cause 1):", np.round(got1, 4), "期望", exp1, "->", "通过" if ok1 else "失败")
    print("手算例 CIF(cause 2):", np.round(got2, 4), "期望", exp2, "->", "通过" if ok2 else "失败")
    tot = got1[-1] + got2[-1]
    print(f"两因累积之和 = {tot:.4f}（无删失残留时应为 1.0）->",
          "通过" if np.isclose(tot, 1.0) else "失败")
    return ok1 and ok2 and np.isclose(tot, 1.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--endpoint", choices=["midpoint", "right"], default="midpoint",
                    help="区间删失取中点(默认)或首次阴性片时刻(上界)")
    ap.add_argument("--xmax", type=float, default=72,
                    help="横轴上限(h)，默认 72——48 h 后风险集过小，再往右不可靠")
    ap.add_argument("--boot", type=int, default=N_BOOT, help="bootstrap 次数")
    ap.add_argument("--selftest", action="store_true", help="仅运行估计量自检")
    args = ap.parse_args()

    if args.selftest:
        sys.exit(0 if selftest() else 1)

    print("估计量自检:")
    if not selftest():
        sys.exit("自检未通过，终止")
    print()

    df, is_verified = build_events(args.endpoint)
    os.makedirs(OUTDIR, exist_ok=True)
    df.to_csv(os.path.join(OUTDIR, "fig2_events.csv"), index=False, encoding="utf-8-sig")

    n_unstrat = int(df["stratum"].isna().sum())
    d = df[df["stratum"].notna()].copy()
    grid = np.arange(0, args.xmax + 0.5, 0.5)

    log = []
    if not is_verified:
        log.append("!! 未经人工核对，仅供流程验证，不可用于投稿")
    log.append(f"区间删失处理: {args.endpoint}")
    log.append(f"纳入 {len(df)} 例；可分层 {len(d)}，首片无法定位而未分层 {n_unstrat}")
    log.append(f"事件构成: 自然排出 {int((df['event']==PASSAGE).sum())}，"
               f"取出(竞争) {int((df['event']==REMOVAL).sum())}，"
               f"删失 {int((df['event']==CENSORED).sum())}"
               f" ({100*(df['event']==CENSORED).mean():.0f}%)")
    iv = df["interval_h"].dropna()
    if len(iv):
        log.append(f"区间宽度(h): median={iv.median():.1f} "
                   f"IQR={iv.quantile(.25):.1f}–{iv.quantile(.75):.1f} max={iv.max():.1f}")

    log.append("\n累积发生率 (%, 95% CI) —— 自然排出:")
    rows = []
    for s in d["stratum"].unique():
        sub = d[d["stratum"] == s]
        cif = aalen_johansen(sub["time_h"], sub["event"], PASSAGE, np.array(REPORT_TIMES, float))
        lo, hi = bootstrap_ci(sub["time_h"], sub["event"], PASSAGE,
                              np.array(REPORT_TIMES, float), n_boot=args.boot)
        log.append(f"  {s} (n={len(sub)})")
        for i, t in enumerate(REPORT_TIMES):
            log.append(f"    {t:>4.0f} h: {100*cif[i]:5.1f}  ({100*lo[i]:.1f}–{100*hi[i]:.1f})")
            rows.append({"stratum": s, "hours": t, "cif_pct": 100 * cif[i],
                         "lo_pct": 100 * lo[i], "hi_pct": 100 * hi[i]})
    pd.DataFrame(rows).to_csv(os.path.join(OUTDIR, "fig2_cif_table.csv"),
                              index=False, encoding="utf-8-sig")

    pt = permutation_test(d, PASSAGE, grid)
    if pt:
        log.append(f"\n组间差异（置换检验，非 Gray 检验）: "
                   f"统计量={pt[0]:.2f}, P={pt[1]:.3f}")

    plot_figure(d, grid, args)
    txt = "\n".join(log)
    with open(os.path.join(OUTDIR, "fig2_log.txt"), "w", encoding="utf-8") as f:
        f.write(txt + "\n")
    print(txt)
    print(f"\n已写出: {FIGDIR}/fig2_competing_risks.svg / .png, "
          f"{OUTDIR}/fig2_cif_table.csv, fig2_events.csv, fig2_log.txt")


if __name__ == "__main__":
    main()

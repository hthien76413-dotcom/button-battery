#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Table 1 生成：已过食管纽扣电池队列的基线特征 / Baseline characteristics.

输出:
  output/table1.csv           投稿用表（UTF-8-SIG，Excel 可直接打开）
  output/table1.md            速读用
  output/table1_log.txt       口径、分组定义、缺失情况、警告
  output/symptom_audit.csv    每一处症状词匹配及其肯定/否定判定与原文片段
  output/analysis_dataset.csv 逐例分析数据集（含患者层面数据，勿提交仓库）

用法:
  python3 scripts/table1.py                     # 默认：内镜 vs 期待治疗
  python3 scripts/table1.py --grouping timing   # 早期内镜/延迟内镜/期待治疗
  python3 scripts/table1.py --no-p              # 不输出 P 值（见下方说明）

关于 Table 1 的 P 值:
  观察性研究的基线表里，P 值检验的是"分组为随机分配"这一零假设——而该假设
  按设计就是假的，所以 P 值在这里意义有限，不少期刊已不建议列出。
  本脚本默认同时给出标准化均数差 (SMD)，并以 SMD 作为衡量基线不均衡的主要指标：
  |SMD| > 0.10 通常提示有意义的不均衡。
"""
import argparse
import os
import sys
from datetime import datetime

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cohort as C

try:
    from scipy import stats
    HAVE_SCIPY = True
except ImportError:
    HAVE_SCIPY = False

OUTDIR = "output"
EARLY_ENDOSCOPY_H = 12      # 早期内镜阈值；方案 §6.4 的敏感性分析改为 24


# ============================================================
# 数据集构建
# ============================================================
def build_dataset(restrict_beyond_esophagus=True):
    xl = C.load_workbook()
    base = xl.parse("病案首页基本信息")
    adm = xl.parse("儿科入院记录")
    xray = xl.parse("X线报告")
    op = xl.parse("住院病历手术记录")
    anes = xl.parse("手麻系统信息")

    core = C.build_core_text(adm)
    warnings = []

    verified, is_verified = C.load_verified_cohort()
    if is_verified:
        battery = verified
        warnings.append(f"使用人工核对队列 {C.VERIFIED_LIST}（{len(battery)} 例）")
    else:
        battery = C.ascertain(core, C.BATTERY_KW)
        warnings.append(
            "!! 未找到人工核对队列，回退到自动关键词筛查。\n"
            "!! 关键词筛查已证实存在假阳性（有一例术中诊断实为磁珠致小肠穿孔）。\n"
            f"!! 本表仅供流程验证，不可用于投稿。请先完成核对并写入 {C.VERIFIED_LIST}。"
        )

    # ---- 首次 X 线定位 ----
    xray = xray.copy()
    xray["_t"] = pd.to_datetime(xray["检查时间"], errors="coerce")
    bx = xray[xray["科研就诊编号"].isin(battery)]
    first = bx.sort_values("_t").groupby("科研就诊编号").first()
    nfilms = bx.groupby("科研就诊编号").size()

    # ---- 手术/内镜时间：优先手麻系统（含时分），手术记录仅到日期 ----
    anes = anes.copy()
    anes["_start"] = pd.to_datetime(anes["手术开始时间"], errors="coerce")
    first_proc = (
        anes[anes["科研就诊编号"].isin(battery)]
        .sort_values("_start").groupby("科研就诊编号")["_start"].first()
    )
    bo = op[op["科研就诊编号"].isin(battery)].copy()
    bo["proc"] = bo["手术名称"].map(C.classify_procedure)
    proc_kinds = bo.groupby("科研就诊编号")["proc"].apply(lambda s: ";".join(sorted(set(s))))

    base = base.copy()
    base["_admit"] = pd.to_datetime(base["入院日期"], errors="coerce")
    adm_idx = adm.set_index("科研就诊编号")
    b = base[base["科研就诊编号"].isin(battery)].copy()

    rows = []
    for _, r in b.iterrows():
        v = r["科研就诊编号"]
        film = first.loc[v] if v in first.index else None
        film_txt = (str(film["检查所见"]) + " " + str(film["检查结论"])) if film is not None else ""
        loc = C.classify_location(film_txt) if film_txt.strip() else "no_film"

        proc = proc_kinds.get(v, None)
        t_proc = first_proc.get(v, pd.NaT)
        hrs_to_proc = (
            (t_proc - r["_admit"]).total_seconds() / 3600
            if pd.notna(t_proc) and pd.notna(r["_admit"]) else None
        )

        txt = core.get(v, "")
        flags = C.symptom_flags(txt)
        a = adm_idx.loc[v] if v in adm_idx.index else None

        rows.append({
            "科研就诊编号": v,
            "科研患者编号": r["科研患者编号"],
            "age_y": r["年龄（岁）"],
            "sex": r["性别"],
            "weight_kg": (a["体重(kg)"] if a is not None else None),
            "admit_year": r["_admit"].year if pd.notna(r["_admit"]) else None,
            "los_days": r["实际住院天数"],
            "delay_h": C.parse_delay_hours(a["主诉"] if a is not None else ""),
            "first_film_location": loc,
            "n_films": int(nfilms.get(v, 0)),
            "battery_mm": C.parse_battery_size(txt),
            "multiple_objects": bool(
                any(k in txt for k in ["两枚", "2枚", "三枚", "3枚", "多枚", "数枚", "两颗", "2颗"])
            ),
            "procedure_kinds": proc,
            "hours_to_procedure": hrs_to_proc,
            **flags,
        })

    df = pd.DataFrame(rows)
    df["any_symptom"] = df[list(C.SYMPTOMS)].any(axis=1)

    # ---- 派生分类变量 ----
    df["age_group"] = pd.cut(
        df["age_y"], bins=[-np.inf, 2, 5, np.inf],
        labels=["<2 岁", "2–<5 岁", "≥5 岁"], right=False,
    )
    df["delay_group"] = np.where(
        df["delay_h"].isna(), "未记录",
        np.where(df["delay_h"] <= 12, "≤12 h", ">12 h"),
    )
    df["size_group"] = np.where(
        df["battery_mm"].isna(), "未记录",
        np.where(df["battery_mm"] >= 20, "≥20 mm", "<20 mm"),
    )
    df["era"] = np.where(df["admit_year"] <= 2020, "2016–2020", "2021–2026")
    df["location_group"] = df["first_film_location"].map(
        lambda x: {"stomach": "胃/十二指肠", "duodenum": "胃/十二指肠"}.get(
            x, "十二指肠以远" if x in ("small_bowel", "colorectal", "passed_or_absent")
            else ("食管" if x == "esophagus" else "未定位/无片")
        )
    )

    # ---- 分组 ----
    df["had_procedure"] = df["procedure_kinds"].notna()
    df["group_binary"] = np.where(df["had_procedure"], "内镜/手术", "期待治疗")
    df["group_timing"] = np.where(
        ~df["had_procedure"], "期待治疗",
        np.where(
            df["hours_to_procedure"].isna(), "内镜(时间未知)",
            np.where(df["hours_to_procedure"] <= EARLY_ENDOSCOPY_H,
                     f"早期内镜(≤{EARLY_ENDOSCOPY_H}h)", f"延迟内镜(>{EARLY_ENDOSCOPY_H}h)"),
        ),
    )

    n_all = len(df)
    if restrict_beyond_esophagus:
        excl_eso = (df["first_film_location"] == "esophagus").sum()
        excl_nofilm = (df["first_film_location"] == "no_film").sum()
        df = df[~df["first_film_location"].isin(["esophagus", "no_film"])].copy()
        warnings.append(
            f"研究人群限定为已过食管：{n_all} 例中排除食管内 {excl_eso} 例、"
            f"无首次影像 {excl_nofilm} 例，余 {len(df)} 例"
        )

    return df, core, battery, warnings


# ============================================================
# 统计
# ============================================================
def smd_continuous(a, b):
    a, b = np.asarray(a, float), np.asarray(b, float)
    a, b = a[~np.isnan(a)], b[~np.isnan(b)]
    if len(a) < 2 or len(b) < 2:
        return None
    pooled = (np.var(a, ddof=1) + np.var(b, ddof=1)) / 2
    if pooled <= 0:
        return 0.0
    return (a.mean() - b.mean()) / np.sqrt(pooled)


def smd_binary(p1, p2):
    denom = (p1 * (1 - p1) + p2 * (1 - p2)) / 2
    if denom <= 0:
        return 0.0
    return (p1 - p2) / np.sqrt(denom)


def fmt_median(s):
    s = pd.to_numeric(s, errors="coerce").dropna()
    if s.empty:
        return "—"
    return f"{s.median():.1f} ({s.quantile(.25):.1f}–{s.quantile(.75):.1f})"


def fmt_np(n, total):
    return "—" if total == 0 else f"{n} ({100 * n / total:.1f})"


def row_continuous(df, col, label, groups, gcol, show_p):
    out = {"变量": f"{label}，中位数 (IQR)", "全体": fmt_median(df[col])}
    vals = []
    for g in groups:
        sub = df.loc[df[gcol] == g, col]
        out[g] = fmt_median(sub)
        vals.append(pd.to_numeric(sub, errors="coerce").dropna())
    out["缺失"] = int(pd.to_numeric(df[col], errors="coerce").isna().sum())
    out["SMD"] = ""
    out["P"] = ""
    if len(groups) == 2:
        d = smd_continuous(vals[0], vals[1])
        out["SMD"] = f"{d:.2f}" if d is not None else ""
        if show_p and HAVE_SCIPY and len(vals[0]) > 1 and len(vals[1]) > 1:
            try:
                out["P"] = f"{stats.mannwhitneyu(vals[0], vals[1], alternative='two-sided').pvalue:.3f}"
            except ValueError:
                pass
    return out


def rows_categorical(df, col, label, groups, gcol, show_p, levels=None):
    """分类变量：一行标题 + 每个水平一行。"""
    if levels is None:   # 未显式指定时才排序；显式给定的顺序即为临床顺序
        levels = sorted(pd.Series(df[col].astype(str)).dropna().unique(), key=str)
    rows = [{"变量": f"{label}，n (%)", "全体": "", **{g: "" for g in groups},
             "缺失": int(df[col].isna().sum()), "SMD": "", "P": ""}]

    disp = df[col].astype(str).replace({"True": "是", "False": "否"})
    df = df.assign(**{col: disp})
    levels = ["是" if l == "True" else ("否" if l == "False" else l) for l in levels]
    tab = pd.crosstab(df[col].astype(str), df[gcol])
    for g in groups:
        if g not in tab.columns:
            tab[g] = 0

    for lv in levels:
        r = {"变量": f"　{lv}", "全体": fmt_np(int((df[col].astype(str) == lv).sum()), len(df))}
        props = []
        for g in groups:
            sub = df[df[gcol] == g]
            n = int((sub[col].astype(str) == lv).sum())
            r[g] = fmt_np(n, len(sub))
            props.append(n / len(sub) if len(sub) else 0.0)
        r["缺失"] = ""
        r["SMD"] = f"{smd_binary(props[0], props[1]):.2f}" if len(groups) == 2 else ""
        r["P"] = ""
        rows.append(r)

    if show_p and HAVE_SCIPY and len(levels) > 1:
        sub_tab = tab.loc[[l for l in levels if l in tab.index], groups]
        if sub_tab.shape[0] > 1 and sub_tab.values.sum() > 0:
            try:
                if sub_tab.shape == (2, 2):
                    p = stats.fisher_exact(sub_tab.values)[1]
                else:
                    # 稀疏格多，优先 Fisher；过大时退回卡方并在日志中标注
                    try:
                        p = stats.fisher_exact(sub_tab.values)[1]
                    except (ValueError, MemoryError):
                        p = stats.chi2_contingency(sub_tab.values)[1]
                rows[0]["P"] = f"{p:.3f}"
            except ValueError:
                pass
    return rows


def build_table1(df, gcol, show_p):
    groups = [g for g in df[gcol].dropna().unique()]
    groups = sorted(groups, key=lambda g: (g == "期待治疗", str(g)))
    rows = []

    rows.append(row_continuous(df, "age_y", "年龄（岁）", groups, gcol, show_p))
    rows += rows_categorical(df, "age_group", "年龄分组", groups, gcol, show_p,
                             levels=["<2 岁", "2–<5 岁", "≥5 岁"])
    rows += rows_categorical(df, "sex", "性别", groups, gcol, show_p)
    rows.append(row_continuous(df, "weight_kg", "体重（kg）", groups, gcol, show_p))
    rows.append(row_continuous(df, "delay_h", "误食至就诊（h）", groups, gcol, show_p))
    rows += rows_categorical(df, "delay_group", "就诊延迟分组", groups, gcol, show_p,
                             levels=["≤12 h", ">12 h", "未记录"])
    rows += rows_categorical(df, "size_group", "电池直径", groups, gcol, show_p,
                             levels=["<20 mm", "≥20 mm", "未记录"])
    rows += rows_categorical(df, "multiple_objects", "多枚异物", groups, gcol, show_p,
                             levels=["是"])
    rows += rows_categorical(df, "location_group", "首次影像定位", groups, gcol, show_p,
                             levels=["胃/十二指肠", "十二指肠以远", "未定位/无片"])
    rows += rows_categorical(df, "any_symptom", "就诊时有症状", groups, gcol, show_p,
                             levels=["是"])
    for key, label in [("vomiting", "　恶心/呕吐"), ("pain", "　腹痛/胸痛"),
                       ("fever", "　发热"), ("hematemesis_melena", "　呕血/黑便"),
                       ("dysphagia_drooling", "　吞咽困难/流涎"),
                       ("feeding_refusal", "　拒食/纳差")]:
        r = {"变量": f"{label}，n (%)",
             "全体": fmt_np(int(df[key].sum()), len(df)), "缺失": "", "SMD": "", "P": ""}
        props = []
        for g in groups:
            sub = df[df[gcol] == g]
            r[g] = fmt_np(int(sub[key].sum()), len(sub))
            props.append(sub[key].mean() if len(sub) else 0.0)
        if len(groups) == 2:
            r["SMD"] = f"{smd_binary(props[0], props[1]):.2f}"
        rows.append(r)
    rows.append(row_continuous(df, "n_films", "X 线检查次数", groups, gcol, show_p))
    rows.append(row_continuous(df, "los_days", "住院天数", groups, gcol, show_p))
    rows += rows_categorical(df, "era", "入院时期", groups, gcol, show_p,
                             levels=["2016–2020", "2021–2026"])

    cols = ["变量", "全体"] + groups + ["缺失", "SMD", "P"]
    if not show_p:
        cols.remove("P")
        for r in rows:
            r.pop("P", None)
    t = pd.DataFrame(rows)
    for c in cols:
        if c not in t.columns:
            t[c] = ""
    return t[cols].fillna(""), groups


# ============================================================
# 主流程
# ============================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--grouping", choices=["binary", "timing"], default="binary",
                    help="binary: 内镜/手术 vs 期待治疗（默认）；timing: 按内镜时机再分")
    ap.add_argument("--no-p", action="store_true", help="不输出 P 值，仅用 SMD")
    ap.add_argument("--all-locations", action="store_true",
                    help="不限定已过食管，纳入全部电池病例")
    args = ap.parse_args()

    os.makedirs(OUTDIR, exist_ok=True)
    show_p = not args.no_p
    gcol = "group_binary" if args.grouping == "binary" else "group_timing"

    df, core, battery, warnings = build_dataset(
        restrict_beyond_esophagus=not args.all_locations
    )
    table, groups = build_table1(df, gcol, show_p)

    # ---- 输出 ----
    table.to_csv(os.path.join(OUTDIR, "table1.csv"), index=False, encoding="utf-8-sig")
    df.to_csv(os.path.join(OUTDIR, "analysis_dataset.csv"), index=False, encoding="utf-8-sig")

    audit = []
    for v in df["科研就诊编号"]:
        for h in C.extract_symptoms(core.get(v, "")):
            audit.append({"科研就诊编号": v, **h})
    pd.DataFrame(audit).to_csv(
        os.path.join(OUTDIR, "symptom_audit.csv"), index=False, encoding="utf-8-sig"
    )

    header = f"Table 1　基线特征，按处置分组（n = {len(df)}）"
    md = [f"# {header}", ""]
    if not os.path.exists(C.VERIFIED_LIST):
        md += ["> **未经人工核对——仅供流程验证，不可用于投稿。**", ""]
    md.append("| " + " | ".join(table.columns) + " |")
    md.append("|" + "|".join(["---"] * len(table.columns)) + "|")
    for _, r in table.iterrows():
        md.append("| " + " | ".join(str(x) for x in r) + " |")
    md += ["", "SMD = 标准化均数差；|SMD| > 0.10 提示基线不均衡。",
           "连续变量 Mann-Whitney U 检验，分类变量 Fisher 精确检验。"]
    with open(os.path.join(OUTDIR, "table1.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md) + "\n")

    log = [f"生成时间: {datetime.now():%Y-%m-%d %H:%M}",
           f"分组变量: {gcol}   分组: {groups}",
           f"早期内镜阈值: {EARLY_ENDOSCOPY_H} h",
           f"scipy: {'可用' if HAVE_SCIPY else '不可用——已跳过 P 值'}",
           f"纳入分析: {len(df)} 例", ""]
    log += ["警告与口径:"] + [f"  {w}" for w in warnings] + [""]
    log.append("各变量缺失情况:")
    for c in ["age_y", "sex", "weight_kg", "delay_h", "battery_mm", "los_days"]:
        log.append(f"  {c:16s} 缺失 {int(df[c].isna().sum()):3d} / {len(df)}")
    log += ["", "分组构成:"]
    for g, n in df[gcol].value_counts().items():
        log.append(f"  {g}: {n}")
    log += ["", "症状提取: 已做否定识别（见 cohort.extract_symptoms）。",
            "  现病史几乎全为否定式套话，不做否定识别会把整个队列判成有症状。",
            "  output/symptom_audit.csv 列出每一处匹配及判定，请人工抽查。"]
    with open(os.path.join(OUTDIR, "table1_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n")

    print("\n".join(log))
    print("\n" + "=" * 78)
    print(header)
    print("=" * 78)
    print(table.to_string(index=False))
    print(f"\n已写出: {OUTDIR}/table1.csv, table1.md, table1_log.txt, "
          f"symptom_audit.csv, analysis_dataset.csv")


if __name__ == "__main__":
    main()

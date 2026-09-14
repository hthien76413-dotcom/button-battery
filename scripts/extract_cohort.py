#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纽扣电池消化道异物队列提取 / Button battery cohort extraction

从 10 年消化道异物原始数据中筛出纽扣电池病例，输出:
  output/screening_list.csv   逐例筛查表（供人工核对，STROBE 流程图用）
  output/feasibility.txt      可行性汇总数字

重要: 关键词筛选存在假阳性（例如病历文本提及"电池"但术中诊断为磁珠），
      screening_list.csv 中的 need_manual_review 列标记需人工核对的病例。
      任何发表前的最终队列必须经病历人工确认。

用法: python3 scripts/extract_cohort.py
"""
import os
import re
import sys
from collections import Counter

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cohort as C

XLSX = C.XLSX
OUTDIR = "output"

# 病例判定、部位映射、时间解析、操作分类等规则统一来自 cohort.py，
# 避免同一套规则在两个脚本里各写一份（方案要求这些规则作为补充材料公开）。
BATTERY_KW = C.BATTERY_KW
MAGNET_KW = C.MAGNET_KW
COIN_KW = C.COIN_KW
LOC_RULES = C.LOC_RULES
collect_text = C.collect_text
parse_delay_hours = C.parse_delay_hours
classify_location = C.classify_location
classify_procedure = C.classify_procedure


def main():
    if not os.path.exists(XLSX):
        sys.exit(f"未找到数据文件: {XLSX}")
    os.makedirs(OUTDIR, exist_ok=True)
    xl = C.load_workbook(XLSX)

    base = xl.parse("病案首页基本信息")
    adm = xl.parse("儿科入院记录")
    xray = xl.parse("X线报告")
    op = xl.parse("住院病历手术记录")

    # ---- 病例判定只用"主诉+现病史+初步诊断"，避开既往史否定句污染 ----
    adm = adm.copy()
    adm["core"] = (
        adm["主诉"].fillna("") + " " + adm["现病史"].fillna("") + " " + adm["初步诊断"].fillna("")
    )
    core = dict(zip(adm["科研就诊编号"], adm["core"]))

    def has(t, kws):
        return any(k in t for k in kws)

    battery = {v for v, t in core.items() if has(t, BATTERY_KW)}
    coin = {v for v, t in core.items() if has(t, COIN_KW) and v not in battery}
    magnet = {v for v, t in core.items() if has(t, MAGNET_KW)}

    # ---- 全病历文本（含手术记录/出院记录），用于假阳性侦测 ----
    alltext = collect_text(xl, xl.sheet_names)

    # ---- 首次 X 线定位 ----
    xray = xray.copy()
    xray["t"] = pd.to_datetime(xray["检查时间"], errors="coerce")
    first_film = (
        xray[xray["科研就诊编号"].isin(battery)]
        .sort_values("t")
        .groupby("科研就诊编号")
        .first()
    )
    nfilms = xray[xray["科研就诊编号"].isin(battery)].groupby("科研就诊编号").size()
    # 自然排出需检索全部复查片，而非仅首张
    filmtexts = (
        xray[xray["科研就诊编号"].isin(battery)]
        .assign(_t=lambda d: d["检查所见"].astype(str) + "。" + d["检查结论"].astype(str))
        .groupby("科研就诊编号")["_t"]
        .apply(list)
    )

    # ---- 操作 ----
    bo = op[op["科研就诊编号"].isin(battery)].copy()
    bo["proc"] = bo["手术名称"].map(classify_procedure)
    proc_by_visit = bo.groupby("科研就诊编号")["proc"].apply(lambda s: ";".join(sorted(set(s))))
    optext = bo.groupby("科研就诊编号")["手术经过"].apply(
        lambda s: " ".join(str(x) for x in s if pd.notna(x))
    )

    # ---- 逐例筛查表 ----
    b = base[base["科研就诊编号"].isin(battery)].copy()
    cc = dict(zip(adm["科研就诊编号"], adm["主诉"]))
    rows = []
    for _, r in b.iterrows():
        v = r["科研就诊编号"]
        ff = first_film.loc[v] if v in first_film.index else None
        film_txt = (
            (str(ff["检查所见"]) + " " + str(ff["检查结论"])) if ff is not None else ""
        )
        ot = optext.get(v, "")
        at = alltext.get(v, "")
        rows.append({
            "科研就诊编号": v,
            "科研患者编号": r["科研患者编号"],
            "第几次住院": r["第几次住院"],
            "sex": r["性别"],
            "age_y": r["年龄（岁）"],
            "admit_year": pd.to_datetime(r["入院日期"]).year,
            "los_days": r["实际住院天数"],
            "delay_h": parse_delay_hours(cc.get(v, "")),
            "first_film_location": classify_location(film_txt) if film_txt else "no_film",
            "n_films": int(nfilms.get(v, 0)),
            # 逐片否定感知判定；旧的合并文本正则会把
            # "未见液平面。消化道异物，位置较前" 误判为已排出
            "documented_passage": any(
                C.classify_film_state(t)[0] == "absent" for t in filmtexts.get(v, [])
            ),
            "procedure": proc_by_visit.get(v, "none_observation"),
            "multiple_objects": bool(re.search(r"两枚|2枚|三枚|3枚|多枚|数枚|两颗|2颗", core.get(v, ""))),
            "size_documented": bool(re.search(r"\d+\s*mm|CR\d|直径|\d+\s*毫米", core.get(v, ""))),
            "esoph_mucosal_injury_in_op": bool(re.search(r"食管.{0,25}(溃疡|糜烂|灼|发白|坏死)", ot)),
            "gastric_mucosal_injury_in_op": bool(re.search(r"胃.{0,25}(溃疡|糜烂|灼伤|坏死)", ot)),
            "futile_endoscopy": bool(re.search(r"未见异物|未找到异物|反复寻找未见", ot)),
            # 假阳性信号: 全文提及磁性异物，或术中诊断未提电池
            "need_manual_review": bool(has(at, MAGNET_KW)) or bool(
                re.search(r"磁珠|巴克球", str(bo[bo["科研就诊编号"] == v]["术中诊断"].to_list()))
            ),
        })
    sl = pd.DataFrame(rows).sort_values(["admit_year", "科研就诊编号"])
    sl.to_csv(os.path.join(OUTDIR, "screening_list.csv"), index=False, encoding="utf-8-sig")

    # ---- 可行性汇总 ----
    L = []
    L.append("=" * 66)
    L.append("纽扣电池队列可行性汇总 (自动生成, 需人工核对后定稿)")
    L.append("=" * 66)
    L.append(f"数据期间: {pd.to_datetime(base['入院日期']).min():%Y-%m-%d} ~ {pd.to_datetime(base['入院日期']).max():%Y-%m-%d}")
    L.append(f"消化道异物总就诊: {len(base)} (患者 {base['科研患者编号'].nunique()})")
    L.append("")
    L.append(f"纽扣电池: {len(battery)} 就诊 / {b['科研患者编号'].nunique()} 患者")
    L.append(f"  同期硬币(拟作惰性异物对照): {len(coin)}")
    L.append(f"  同期磁性异物(姊妹课题): {len(magnet)}")
    L.append(f"  电池与磁性共摄入(需排除): {len(battery & magnet)}")
    L.append("")
    L.append(f"年龄(岁) median={b['年龄（岁）'].median():.1f} IQR={b['年龄（岁）'].quantile(.25):.1f}-{b['年龄（岁）'].quantile(.75):.1f}")
    L.append(f"男性 {(b['性别']=='男性').sum()} ({100*(b['性别']=='男性').mean():.0f}%)")
    L.append(f"住院天数 median={b['实际住院天数'].median():.0f}")
    d = sl["delay_h"].dropna()
    L.append(f"误食至就诊(h) median={d.median():.1f} IQR={d.quantile(.25):.1f}-{d.quantile(.75):.1f}; >12h {(d>12).sum()} 例")
    L.append("")
    L.append("首次X线定位:")
    for k, n in sl["first_film_location"].value_counts().items():
        L.append(f"  {k}: {n}")
    beyond = (~sl["first_film_location"].isin(["esophagus", "no_film"])).sum()
    L.append(f"  => 已过食管(研究人群): 约 {beyond}")
    L.append("")
    L.append("处置分组:")
    for k, n in sl["procedure"].value_counts().items():
        L.append(f"  {k}: {n}")
    L.append("")
    L.append("结局信号(文本挖掘, 需人工分级):")
    L.append(f"  操作记录描述食管黏膜损伤: {sl['esoph_mucosal_injury_in_op'].sum()}")
    L.append(f"  操作记录描述胃黏膜损伤: {sl['gastric_mucosal_injury_in_op'].sum()}")
    L.append(f"  进镜未找到异物(无效内镜): {sl['futile_endoscopy'].sum()}")
    L.append(f"  影像记录自然排出: {sl['documented_passage'].sum()}")
    L.append(f"  记录了电池规格(mm/CR/直径): {sl['size_documented'].sum()} <-- 主要数据缺失")
    L.append(f"  需人工核对(假阳性风险): {sl['need_manual_review'].sum()}")
    L.append("")
    L.append("X线复查次数分布: " + str(dict(sorted(Counter(sl["n_films"]).items()))))
    txt = "\n".join(L)
    with open(os.path.join(OUTDIR, "feasibility.txt"), "w", encoding="utf-8") as f:
        f.write(txt + "\n")
    print(txt)
    print(f"\n已写出: {OUTDIR}/screening_list.csv ({len(sl)} 行), {OUTDIR}/feasibility.txt")


if __name__ == "__main__":
    main()

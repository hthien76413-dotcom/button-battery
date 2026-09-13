#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纽扣电池研究的共享队列逻辑 / Shared cohort logic.

extract_cohort.py 与 table1.py 都从这里取判定规则，确保部位映射、病例判定、
症状提取三套规则只有一个来源——方案中这些规则要作为补充材料公开，
不能在两个脚本里各写一份。
"""
import os
import re

import pandas as pd

XLSX = "10年消化道异物原始数据.xlsx"

# 人工核对后的最终队列。存在时优先使用，否则回退到自动筛查（并在输出中警告）。
VERIFIED_LIST = "data/verified_cohort.csv"

# ---------------- 病例判定关键词 ----------------
BATTERY_KW = ["纽扣电池", "扣状电池", "钮扣电池", "纽扣型电池", "扣式电池"]
MAGNET_KW = ["磁珠", "磁铁", "巴克球", "吸铁石", "磁力球", "磁性异物"]
COIN_KW = ["硬币"]

# ---------------- 首次 X 线定位映射规则 ----------------
# 放射报告常只写解剖分区而非脏器，故按分区推断。须两名读片者独立复核并报告 κ。
LOC_RULES = [
    ("esophagus", r"食管|食道|颈部|胸[段部]|纵隔|气管隆突|T\d+\s*水平"),
    ("stomach", r"胃[内腔区壁]|胃泡|胃底|胃窦|左上腹|中上腹|上腹部?见"),
    ("duodenum", r"十二指肠"),
    ("small_bowel", r"回肠|空肠|小肠"),
    ("colorectal", r"盆腔|下腹|结肠|直肠|骶髂|乙状"),
    ("passed_or_absent", r"未见.{0,8}(不透|异物)|异物.{0,6}(已)?(消失|排出)|阳性异物已排出"),
    ("abdomen_unspecified", r"腹部|腹腔|腰椎|L\d+\s*水平|中线|右中腹|左中腹"),
]

# ---------------- 症状词表 ----------------
# 键为报告用名，值为匹配词。不收 "食欲""精神" 等——出院套话 "精神食欲睡眠一般"
# 会造成大量假阳性。
SYMPTOMS = {
    "vomiting": ["呕吐", "恶心"],
    "hematemesis_melena": ["呕血", "黑便", "血便", "便血"],
    "dysphagia_drooling": ["吞咽困难", "流涎", "流口水", "咽下困难", "拒吞"],
    "pain": ["腹痛", "胸痛", "腹部疼痛"],
    "fever": ["发热", "发烧"],
    "feeding_refusal": ["拒食", "纳差", "拒奶", "进食减少"],
}

# 否定与肯定标记。刻意不收单字 "不"——"不慎误食" 是固定搭配，会误伤。
NEG_MARKERS = ["无", "未", "否认", "没有", "不伴", "阴性"]
POS_MARKERS = ["伴有", "伴", "出现", "可见", "有", "阳性", "诉"]
SENT_SEP = r"[。；;\n\r]"
FRAG_SEP = r"[、,，]"


# ================= 载入 =================
def load_workbook(path=XLSX):
    if not os.path.exists(path):
        raise SystemExit(f"未找到数据文件: {path}")
    return pd.ExcelFile(path)


def collect_text(xl, sheets, keep_cols=None):
    """按就诊编号聚合指定 sheet 的文本列。"""
    out = {}
    for s in sheets:
        df = xl.parse(s)
        if "科研就诊编号" not in df.columns:
            continue
        cols = [
            c for c in df.columns
            if c not in ("科研患者编号", "科研就诊编号")
            and (pd.api.types.is_string_dtype(df[c]) or df[c].dtype == object)
        ]
        if keep_cols:
            cols = [c for c in cols if c in keep_cols]
        for _, r in df.iterrows():
            v = r["科研就诊编号"]
            out[v] = out.get(v, "") + " " + " ".join(
                str(r[c]) for c in cols if pd.notna(r[c])
            )
    return out


def build_core_text(adm):
    """病例判定只用主诉 + 现病史 + 初步诊断。

    既往史/家族史充斥 "无输血史""否认出血性疾病" 一类否定句，
    纳入会严重污染关键词计数。
    """
    core = (
        adm["主诉"].fillna("") + " "
        + adm["现病史"].fillna("") + " "
        + adm["初步诊断"].fillna("")
    )
    return dict(zip(adm["科研就诊编号"], core))


def ascertain(core_text, keywords):
    return {v for v, t in core_text.items() if any(k in t for k in keywords)}


def load_verified_cohort(path=VERIFIED_LIST):
    """载入人工核对后的队列。

    返回 (就诊编号集合, 是否为已核对队列)。CSV 需含列 `科研就诊编号`；
    若另有 `include` 列，则仅保留其值为 1/true/yes 的行。
    """
    if not os.path.exists(path):
        return None, False
    df = pd.read_csv(path)
    if "科研就诊编号" not in df.columns:
        raise SystemExit(f"{path} 缺少列 `科研就诊编号`")
    if "include" in df.columns:
        keep = df["include"].astype(str).str.strip().str.lower()
        df = df[keep.isin(["1", "true", "yes", "y", "是"])]
    return set(df["科研就诊编号"]), True


# ================= 字段解析 =================
def parse_delay_hours(chief_complaint):
    """从主诉解析误食至就诊时间（小时）。"""
    s = str(chief_complaint)
    m = re.search(r"(\d+(?:\.\d+)?)\s*(小时|天|日|周|月)", s)
    if not m:
        return None
    return float(m.group(1)) * {
        "小时": 1, "天": 24, "日": 24, "周": 168, "月": 720
    }[m.group(2)]


def classify_location(text):
    for label, pat in LOC_RULES:
        if re.search(pat, text):
            return label
    return "undetermined"


BEYOND_ESOPHAGUS = {
    "stomach", "duodenum", "small_bowel", "colorectal",
    "abdomen_unspecified", "passed_or_absent",
}


def classify_procedure(name):
    n = re.sub(r"\s+", "", str(name))
    if re.search(r"剖腹|修补|肠切|腹腔镜", n):
        return "surgery"
    if "食管" in n and re.search(r"取出|去除", n):
        return "esophageal_endoscopic_removal"
    if "十二指肠" in n and re.search(r"取出|去除", n):
        return "duodenal_endoscopic_removal"
    if re.search(r"异物", n) and re.search(r"取出|去除", n):
        return "gastric_endoscopic_removal"
    if re.search(r"胃镜检查|内镜检查", n):
        return "diagnostic_endoscopy_only"
    return "other"


def parse_battery_size(text):
    """从病历文本解析电池直径 (mm)。本队列仅约 3/141 例有记录。"""
    t = str(text)
    m = re.search(r"直径[约为]?\s*(\d+(?:\.\d+)?)\s*(mm|毫米|cm|厘米)", t)
    if not m:
        m = re.search(r"(\d+(?:\.\d+)?)\s*(mm|毫米)", t)
    if m:
        val = float(m.group(1))
        return val * 10 if m.group(2) in ("cm", "厘米") else val
    if re.search(r"CR\s*20\d\d", t):        # CR2016 / CR2025 / CR2032 均为 20 mm
        return 20.0
    return None


# ================= 否定感知的症状提取 =================
def _first_marker(frag, markers):
    idxs = [frag.find(m) for m in markers if m in frag]
    return min(idxs) if idxs else None


def extract_symptoms(text):
    """从自由文本提取症状，并判断每处是肯定还是否定。

    本队列的现病史几乎全是否定式套话（"无发热、无恶心、无呕吐、无腹痛"），
    141 例中有 140 例出现症状词——朴素匹配会把整个队列判成有症状。

    规则：按句切分；句内再按顿号/逗号切成片段；逐片段判定极性，
    并在句内向后传递。这样 "无腹痛、腹胀等不适" 中的 "腹胀" 能继承前一片段
    的否定（单个 "无" 辖域覆盖整个并列），而 "伴呕吐、腹痛" 则两者均为肯定。
    句末重置传递状态。

    返回 [{symptom, term, negated, snippet}, ...]，肯定与否定的匹配全部返回，
    供人工审核（见 output/symptom_audit.csv）。
    """
    t = str(text).replace("不慎", "意外")      # "不慎误食" 是固定搭配，不是否定
    hits = []
    for sentence in re.split(SENT_SEP, t):
        if not sentence.strip():
            continue
        carry_neg = False
        for frag in re.split(FRAG_SEP, sentence):
            if not frag.strip():
                continue
            neg_i = _first_marker(frag, NEG_MARKERS)
            pos_i = _first_marker(frag, POS_MARKERS)
            if neg_i is not None and (pos_i is None or neg_i < pos_i):
                carry_neg = True
            elif pos_i is not None and (neg_i is None or pos_i < neg_i):
                carry_neg = False
            # 两者皆无时沿用上一片段的极性

            for symptom, terms in SYMPTOMS.items():
                for term in terms:
                    if term in frag:
                        negated = carry_neg
                        # 明确的次数描述（"呕吐3次"）推翻继承来的否定
                        if re.search(term + r"\s*\d+\s*(次|余次)", frag):
                            negated = False
                        hits.append({
                            "symptom": symptom,
                            "term": term,
                            "negated": negated,
                            "snippet": frag.strip()[:60],
                        })
                        break
    return hits


def symptom_flags(text):
    """返回 {symptom: bool}，仅计肯定的匹配。"""
    hits = extract_symptoms(text)
    flags = {k: False for k in SYMPTOMS}
    for h in hits:
        if not h["negated"]:
            flags[h["symptom"]] = True
    return flags

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DR/DICOM 标定信息核查与电池直径换算 / DICOM calibration audit & battery sizing.

背景
----
病历里 123/125 例未记录电池直径，而 ≥20 mm 是文献中最强的风险因素 (OR 4.34)，
也是 Fig 3 的一条分层线。该信息其实在已拍的腹平片里——只是从未从图像上读取。
本脚本回答两个问题:

  1. PACS 导出的 DICOM 还带不带标定信息？（没有 PixelSpacing 这条路就走不通）
  2. 放大校正后，≥20 mm 这个分类在合理的放大倍数范围内稳不稳？

投影放大是关键坑：电池在体内离探测器有距离，成像放大约 1.1–1.2 倍，
而分层阈值恰好是 20 mm。直接拿卡尺读数卡 20 mm 会系统性高估。

隐私
----
本脚本**不输出任何患者身份信息**。姓名、病历号、检查号等仅报告"是否存在"，
用于判断是否需要去标识；PatientID 以截断哈希输出，使同一患者的多张片子能
归组，但无法反推原值。

依赖
----
  pip install pydicom

用法
----
  # 1) 核查一批 DICOM 的标定信息
  python3 scripts/dicom_calibration.py audit /path/to/dicom_dir

  # 2) 单例换算：在 PACS 里量得投影直径 22.4 mm，换算真实直径并吸附到标准规格
  python3 scripts/dicom_calibration.py size 22.4 --sid 1000 --sod 880

  # 3) 不知道 SOD 时，用体表到探测器的估计距离（cm）
  python3 scripts/dicom_calibration.py size 22.4 --sid 1000 --odd 10
"""
import argparse
import hashlib
import os
import sys

try:
    import pydicom
    from pydicom.errors import InvalidDicomError
except ImportError:
    sys.exit("需要 pydicom：pip install pydicom")

import pandas as pd

OUTDIR = "output"

# 纽扣/扣式电池的标准直径 (mm)。CR 型号前两位即直径：CR2032 = 20 mm。
STANDARD_DIAMETERS = [4.8, 5.8, 6.8, 7.9, 9.5, 11.6, 12.5, 16.0, 20.0, 23.0, 24.5]
COMMON_NAMES = {
    12.5: "CR1216/1220/1225",
    16.0: "CR1616/1620/1632",
    20.0: "CR2016/2025/2032",
    23.0: "CR2320/2330/2354",
    24.5: "CR2430/2450",
}

# 放大倍数的敏感性范围：物体距探测器 5–20 cm
ODD_RANGE_CM = (5.0, 20.0)
ODD_DEFAULT_CM = 10.0

# 需核查是否存在的身份标签（只报存在与否，不输出值）
PHI_TAGS = [
    ("PatientName", (0x0010, 0x0010)),
    ("PatientID", (0x0010, 0x0020)),
    ("PatientBirthDate", (0x0010, 0x0030)),
    ("AccessionNumber", (0x0008, 0x0050)),
    ("InstitutionName", (0x0008, 0x0080)),
    ("ReferringPhysicianName", (0x0008, 0x0090)),
]


def _get(ds, keyword, default=None):
    v = getattr(ds, keyword, default)
    if v is None or (isinstance(v, str) and not v.strip()):
        return default
    return v


def _first_float(v):
    """PixelSpacing 等为多值 (row, col)，取第一个并转 float。"""
    if v is None:
        return None
    try:
        if isinstance(v, (list, tuple)) or hasattr(v, "__len__") and not isinstance(v, str):
            return float(v[0])
        return float(v)
    except (TypeError, ValueError, IndexError):
        return None


def pseudonymize(value):
    if value is None:
        return None
    return hashlib.sha256(str(value).encode()).hexdigest()[:12]


# ============================================================
# 标定核查
# ============================================================
def read_one(path):
    try:
        ds = pydicom.dcmread(path, stop_before_pixels=True, force=False)
    except (InvalidDicomError, OSError):
        return None

    imager_ps = _first_float(_get(ds, "ImagerPixelSpacing"))
    pixel_ps = _first_float(_get(ds, "PixelSpacing"))
    det_ps = _first_float(_get(ds, "DetectorElementSpacing"))
    sid = _first_float(_get(ds, "DistanceSourceToDetector"))
    sod = _first_float(_get(ds, "DistanceSourceToPatient"))
    mag = _first_float(_get(ds, "EstimatedRadiographicMagnificationFactor"))

    image_type = _get(ds, "ImageType", [])
    image_type = list(image_type) if hasattr(image_type, "__iter__") and not isinstance(image_type, str) else [str(image_type)]

    # 探测器平面的像素间距：投影摄影首选 ImagerPixelSpacing
    spacing = imager_ps or pixel_ps or det_ps
    spacing_source = (
        "ImagerPixelSpacing" if imager_ps else
        "PixelSpacing" if pixel_ps else
        "DetectorElementSpacing" if det_ps else None
    )

    if mag:
        mag_source = "EstimatedRadiographicMagnificationFactor"
    elif sid and sod and sod > 0:
        mag = sid / sod
        mag_source = "SID/SOD"
    else:
        mag_source = None

    return {
        "file": os.path.basename(path),
        "patient_key": pseudonymize(_get(ds, "PatientID")),
        "study_date": _get(ds, "StudyDate"),          # 用于与就诊记录对应
        "modality": _get(ds, "Modality"),
        "body_part": _get(ds, "BodyPartExamined"),
        "view_position": _get(ds, "ViewPosition"),
        "rows": _get(ds, "Rows"),
        "columns": _get(ds, "Columns"),
        "pixel_spacing_mm": spacing,
        "spacing_source": spacing_source,
        "spacing_calibration_type": _get(ds, "PixelSpacingCalibrationType"),
        "sid_mm": sid,
        "sod_mm": sod,
        "magnification": round(mag, 4) if mag else None,
        "magnification_source": mag_source,
        "is_derived": any(t in ("DERIVED", "SECONDARY") for t in image_type),
        "image_type": "/".join(str(t) for t in image_type[:4]),
        "phi_present": ";".join(
            name for name, tag in PHI_TAGS
            if tag in ds and str(ds[tag].value).strip()
        ),
    }


def cmd_audit(args):
    paths = []
    for root, _, files in os.walk(args.path):
        for f in files:
            if f.lower().endswith((".dcm", ".dicom", ".ima")) or "." not in f:
                paths.append(os.path.join(root, f))
    if not paths:
        sys.exit(f"{args.path} 下未找到文件")

    rows = [r for r in (read_one(p) for p in sorted(paths)) if r is not None]
    if not rows:
        sys.exit(f"扫描了 {len(paths)} 个文件，没有一个能作为 DICOM 读取")

    df = pd.DataFrame(rows)
    os.makedirs(OUTDIR, exist_ok=True)
    out = os.path.join(OUTDIR, "dicom_calibration.csv")
    df.to_csv(out, index=False, encoding="utf-8-sig")

    n = len(df)
    has_sp = df["pixel_spacing_mm"].notna().sum()
    has_mag = df["magnification"].notna().sum()
    derived = int(df["is_derived"].sum())
    phi = int((df["phi_present"].fillna("") != "").sum())

    print("=" * 68)
    print(f"DICOM 标定核查：可读取 {n} / 扫描 {len(paths)} 个文件")
    print("=" * 68)
    print(f"模态: {df['modality'].value_counts().to_dict()}")
    print()
    print(f"有像素间距 (可测量):        {has_sp}/{n}" + ("  <== 可行" if has_sp == n else "  <== 部分缺失"))
    if has_sp:
        print(f"  来源: {df['spacing_source'].value_counts().to_dict()}")
        sp = df["pixel_spacing_mm"].dropna()
        print(f"  间距范围: {sp.min():.4f} – {sp.max():.4f} mm/px")
        cal = df["spacing_calibration_type"].dropna()
        if len(cal):
            print(f"  标定类型: {cal.value_counts().to_dict()}")
    print()
    print(f"可直接算放大倍数:           {has_mag}/{n}")
    if has_mag:
        print(f"  来源: {df['magnification_source'].value_counts().to_dict()}")
        mg = df["magnification"].dropna()
        print(f"  倍数范围: {mg.min():.3f} – {mg.max():.3f}")
    else:
        print(f"  SID/SOD 均缺失 —— 需按经验倍数换算，并做阈值敏感性分析")
    print()
    if derived:
        print(f"!! {derived} 个为 DERIVED/SECONDARY 图像，可能已重采样，标定存疑——勿用于测量")
    if phi:
        print(f"!! {phi} 个文件仍含身份标签 ({df['phi_present'].dropna().iloc[0]} 等)")
        print(f"!! 对外共享或投稿补充材料前必须去标识")
    print()
    if has_sp == n and not derived:
        print("结论: 标定完整，PACS 卡尺测量可行。")
    elif has_sp == 0:
        print("结论: 无像素间距——此路不通。改用椎体宽度作内参，或联系 PACS 管理员确认导出设置。")
    else:
        print("结论: 部分可用。仅对有标定且非 DERIVED 的图像测量，其余记为缺失。")
    print(f"\n逐文件结果: {out}")
    return df


# ============================================================
# 直径换算
# ============================================================
def snap_to_standard(mm, tol=1.2):
    """吸附到最近的标准规格；与次近者差距过小则标为不确定。"""
    if mm is None:
        return None, None, True
    ds = sorted(STANDARD_DIAMETERS, key=lambda d: abs(d - mm))
    nearest, second = ds[0], ds[1]
    ambiguous = abs(abs(nearest - mm) - abs(second - mm)) < tol
    return nearest, COMMON_NAMES.get(nearest), ambiguous


def cmd_size(args):
    measured = args.measured_mm
    print("=" * 68)
    print(f"投影直径（PACS 卡尺读数）: {measured:.2f} mm")
    print("=" * 68)

    if args.sod:
        mags = {"SID/SOD": args.sid / args.sod}
    elif args.mag:
        mags = {"指定倍数": args.mag}
    else:
        odd = args.odd * 10.0                      # cm -> mm
        mags = {f"ODD={args.odd:.0f}cm": args.sid / (args.sid - odd)}

    for label, m in mags.items():
        true_mm = measured / m
        nearest, name, amb = snap_to_standard(true_mm)
        print(f"\n放大倍数 {m:.3f} ({label})")
        print(f"  校正后真实直径: {true_mm:.2f} mm")
        print(f"  最近标准规格:   {nearest} mm" + (f"  ({name})" if name else ""))
        if amb:
            print(f"  !! 与次近规格难以区分——建议记为不确定")
        print(f"  分类:           {'≥20 mm' if nearest >= 20 else '<20 mm'}")

    # ---- 阈值稳定性：这才是关键 ----
    lo_odd, hi_odd = ODD_RANGE_CM
    m_lo = args.sid / (args.sid - lo_odd * 10)     # 物体离探测器近 -> 放大小
    m_hi = args.sid / (args.sid - hi_odd * 10)
    t_hi, t_lo = measured / m_lo, measured / m_hi   # 放大小 -> 校正后大
    n_hi = snap_to_standard(t_hi)[0]
    n_lo = snap_to_standard(t_lo)[0]

    print("\n" + "-" * 68)
    print(f"放大倍数敏感性（物体距探测器 {lo_odd:.0f}–{hi_odd:.0f} cm，SID {args.sid:.0f} mm）")
    print(f"  倍数 {m_lo:.3f} – {m_hi:.3f}  ->  真实直径 {t_lo:.2f} – {t_hi:.2f} mm")
    print(f"  吸附规格 {n_lo} – {n_hi} mm")
    if (n_lo >= 20) == (n_hi >= 20):
        print(f"  ≥20 mm 分类在整个范围内稳定: {'≥20 mm' if n_lo >= 20 else '<20 mm'}")
    else:
        print(f"  !! 分类随放大倍数改变（{'<20' if n_lo < 20 else '≥20'} ↔ "
              f"{'≥20' if n_hi >= 20 else '<20'}）")
        print(f"  !! 该例应记为不确定，并纳入主分析的敏感性检验")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("audit", help="核查一批 DICOM 的标定信息")
    a.add_argument("path", help="DICOM 文件所在目录（递归扫描）")
    a.set_defaults(func=cmd_audit)

    s = sub.add_parser("size", help="由投影直径换算真实直径")
    s.add_argument("measured_mm", type=float, help="PACS 中量得的投影直径 (mm)")
    s.add_argument("--sid", type=float, default=1000.0, help="球管到探测器距离 (mm)，默认 1000")
    s.add_argument("--sod", type=float, default=None, help="球管到物体距离 (mm)，有则优先")
    s.add_argument("--mag", type=float, default=None, help="直接指定放大倍数")
    s.add_argument("--odd", type=float, default=ODD_DEFAULT_CM,
                   help=f"物体到探测器距离 (cm)，默认 {ODD_DEFAULT_CM:.0f}")
    s.set_defaults(func=cmd_size)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

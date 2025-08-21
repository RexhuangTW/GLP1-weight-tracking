
import argparse
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib

# --- Chinese font fallback (fix garbled labels) ---
# Try common CJK fonts; matplotlib will use the first available.
matplotlib.rcParams["font.sans-serif"] = [
    "Noto Sans CJK TC", "Noto Sans CJK SC", "Noto Sans CJK JP",
    "Microsoft JhengHei", "PingFang TC", "Heiti TC", "SimHei",
    "WenQuanYi Micro Hei", "Arial Unicode MS", "DejaVu Sans"
]
matplotlib.rcParams["axes.unicode_minus"] = False  # ensure minus sign renders

REQUIRED_LOGICAL = [
    "日期",
    "早上體重 (kg)",
    "晚上體重 (kg)",
    "早上體脂 (%)",
    "晚上體脂 (%)",
]

OPTIONAL_LOGICAL = [
    "藥物劑量 (mg)",
    "副作用紀錄",
    "每日飲水量 (L)",
]

ALIASES = {
    "日期": ["日期", "date", "日期(yyyy-mm-dd)", "時間", "day"],
    "早上體重 (kg)": ["早上體重 (kg)", "早上體重", "am體重", "體重am", "體重-早", "am weight", "morning weight", "morning_weight", "體重(早)","早上體重kg","體重早"],
    "晚上體重 (kg)": ["晚上體重 (kg)", "晚上體重", "pm體重", "體重pm", "體重-晚", "pm weight", "evening weight", "evening_weight", "體重(晚)","晚上體重kg","體重晚"],
    "早上體脂 (%)": ["早上體脂 (%)", "早上體脂", "am體脂", "體脂am", "am body fat", "morning body fat", "morning_bodyfat", "體脂(早)","體脂早","ambodyfat","morningbodyfat"],
    "晚上體脂 (%)": ["晚上體脂 (%)", "晚上體脂", "pm體脂", "體脂pm", "pm body fat", "evening body fat", "evening_bodyfat", "體脂(晚)","體脂晚","pmbodyfat","eveningbodyfat"],
    "藥物劑量 (mg)": ["藥物劑量 (mg)", "藥物劑量", "劑量", "dose", "dosage", "glp1 dosage"],
    "副作用紀錄": ["副作用紀錄", "副作用", "side effects", "side_effects", "notes"],
    "每日飲水量 (L)": ["每日飲水量 (L)", "飲水量", "水量", "water", "daily water (l)", "water_l"],
}

def norm(s: str) -> str:
    s = str(s)
    s = s.strip().lower()
    s = re.sub(r"[()\[\]％%]", "", s)
    s = re.sub(r"\s+|[_\-]+", "", s)
    return s

def build_alias_map():
    m = {}
    for logical, alist in ALIASES.items():
        m[logical] = set(norm(a) for a in alist)
    return m

ALIAS_MAP = build_alias_map()

def ensure_dirs(path):
    os.makedirs(path, exist_ok=True)

def read_daily_log(master_path, sheet_name=None, header_row=0):
    if sheet_name:
        df = pd.read_excel(master_path, sheet_name=sheet_name, header=header_row)
    else:
        try:
            df = pd.read_excel(master_path, sheet_name="Daily Log", header=header_row)
        except Exception:
            df = pd.read_excel(master_path, header=header_row)
    actual_cols = list(df.columns)
    actual_norm = {norm(c): c for c in actual_cols}
    mapping = {}
    for logical in REQUIRED_LOGICAL + OPTIONAL_LOGICAL:
        found = None
        for key_norm, original in actual_norm.items():
            if key_norm in ALIAS_MAP[logical]:
                found = original
                break
        if found:
            mapping[logical] = found
    missing = [col for col in REQUIRED_LOGICAL if col not in mapping]
    if missing:
        msg = [
            "⚠️ 無法從 Excel 映射以下必要欄位：",
            *[f"- {mcol}" for mcol in missing],
            "",
            "目前偵測到的欄位：",
            *[f"- {c}" for c in actual_cols],
            "",
            "解法：",
            "1) 請確認你的欄位名稱是否與下列其中之一相符（可接受別名）："
        ]
        for logical in REQUIRED_LOGICAL:
            msg.append(f"   • {logical} → {sorted(list(ALIAS_MAP[logical]))}")
        msg.append("2) 或使用 --sheet 與 --header-row 指定正確工作表與標題列（0 表示第一列）。")
        raise ValueError("\n".join(msg))
    df_renamed = df.rename(columns={v: k for k, v in mapping.items()})
    keep = [c for c in REQUIRED_LOGICAL + OPTIONAL_LOGICAL if c in df_renamed.columns]
    df_final = df_renamed[keep].copy()
    df_final["日期"] = pd.to_datetime(df_final["日期"])
    return df_final

def assign_custom_week(df, anchor_date):
    d0 = pd.to_datetime(anchor_date).normalize()
    delta_days = (df["日期"].dt.normalize() - d0).dt.days
    week_idx = (delta_days // 7) + 1  # 1-based
    df2 = df.copy()
    df2["WEEK_IDX"] = week_idx
    return df2

def pick_custom_week(df, anchor_date, week_index=None):
    df2 = assign_custom_week(df, anchor_date)
    if week_index is None:
        target = int(df2["WEEK_IDX"].max())
    else:
        target = int(week_index)
    wdf = df2[df2["WEEK_IDX"] == target].copy()
    if wdf.empty:
        raise ValueError(f"在 anchor={anchor_date} 下，找不到第 {target} 週的資料。")
    start_date = pd.to_datetime(anchor_date) + pandas_offset_weeks(target-1)
    end_date = start_date + pd.Timedelta(days=6)
    tag = f"{start_date.year}-CW{target:02d}"
    return wdf, tag, start_date, end_date

def pandas_offset_weeks(n):
    # helper to move by n weeks as Timedelta
    return pd.Timedelta(days=7*n)

def save_weekly_excel(wdf, out_excel_path):
    base_cols = REQUIRED_LOGICAL
    optional = [c for c in OPTIONAL_LOGICAL if c in wdf.columns]
    cols = base_cols + optional
    wdf.loc[:, cols].to_excel(out_excel_path, index=False)

def make_charts(wdf, out_dir, prefix):
    plt.figure(figsize=(8,5))
    plt.plot(wdf["日期"], wdf["早上體重 (kg)"], marker="o", label="早上體重")
    plt.plot(wdf["日期"], wdf["晚上體重 (kg)"], marker="o", label="晚上體重")
    plt.xlabel("日期"); plt.ylabel("體重 (kg)"); plt.title("體重趨勢"); plt.legend(); plt.grid(True)
    plt.xticks(rotation=30)
    weight_png = os.path.join(out_dir, f"{prefix}_weight_trend.png")
    plt.savefig(weight_png, dpi=150, bbox_inches="tight"); plt.close()

    plt.figure(figsize=(8,5))
    plt.plot(wdf["日期"], wdf["早上體脂 (%)"], marker="o", label="早上體脂")
    plt.plot(wdf["日期"], wdf["晚上體脂 (%)"], marker="o", label="晚上體脂")
    plt.xlabel("日期"); plt.ylabel("體脂 (%)"); plt.title("體脂趨勢"); plt.legend(); plt.grid(True)
    plt.xticks(rotation=30)
    bodyfat_png = os.path.join(out_dir, f"{prefix}_bodyfat_trend.png")
    plt.savefig(bodyfat_png, dpi=150, bbox_inches="tight"); plt.close()

    return weight_png, bodyfat_png

def compute_stats(wdf):
    wdf_sorted = wdf.sort_values("日期")
    stats = {
        "period_start": wdf_sorted["日期"].iloc[0].strftime("%Y/%m/%d"),
        "period_end":   wdf_sorted["日期"].iloc[-1].strftime("%Y/%m/%d"),
        "start_weight_am": float(wdf_sorted["早上體重 (kg)"].iloc[0]),
        "end_weight_am":   float(wdf_sorted["早上體重 (kg)"].iloc[-1]),
        "delta_weight_am": float(wdf_sorted["早上體重 (kg)"].iloc[-1]-wdf_sorted["早上體重 (kg)"].iloc[0]),
        "avg_weight_am":   float(wdf_sorted["早上體重 (kg)"].mean()),
        "start_weight_pm": float(wdf_sorted["晚上體重 (kg)"].iloc[0]),
        "end_weight_pm":   float(wdf_sorted["晚上體重 (kg)"].iloc[-1]),
        "delta_weight_pm": float(wdf_sorted["晚上體重 (kg)"].iloc[-1]-wdf_sorted["晚上體重 (kg)"].iloc[0]),
        "avg_weight_pm":   float(wdf_sorted["晚上體重 (kg)"].mean()),
        "avg_weight_all":  float(wdf_sorted[["早上體重 (kg)","晚上體重 (kg)"]].mean().mean()),
        "start_fat_am": float(wdf_sorted["早上體脂 (%)"].iloc[0]),
        "end_fat_am":   float(wdf_sorted["早上體脂 (%)"].iloc[-1]),
        "delta_fat_am": float(wdf_sorted["早上體脂 (%)"].iloc[-1]-wdf_sorted["早上體脂 (%)"].iloc[0]),
        "avg_fat_am":   float(wdf_sorted["早上體脂 (%)"].mean()),
        "start_fat_pm": float(wdf_sorted["晚上體脂 (%)"].iloc[0]),
        "end_fat_pm":   float(wdf_sorted["晚上體脂 (%)"].iloc[-1]),
        "delta_fat_pm": float(wdf_sorted["晚上體脂 (%)"].iloc[-1]-wdf_sorted["晚上體脂 (%)"].iloc[0]),
        "avg_fat_pm":   float(wdf_sorted["晚上體脂 (%)"].mean()),
        "avg_fat_all":  float(wdf_sorted[["早上體脂 (%)","晚上體脂 (%)"]].mean().mean()),
        "days": int(wdf_sorted.shape[0])
    }
    if "每日飲水量 (L)" in wdf_sorted.columns:
        water = wdf_sorted["每日飲水量 (L)"].dropna()
        stats["avg_water"] = float(water.mean()) if not water.empty else None
    else:
        stats["avg_water"] = None
    return stats

def make_markdown(wdf, stats, png_weight, png_bodyfat, out_md_path, week_tag, start_date, end_date):
    tbl = wdf[["日期","早上體重 (kg)","晚上體重 (kg)","早上體脂 (%)","晚上體脂 (%)"]].copy()
    tbl["日期"] = tbl["日期"].dt.strftime("%m/%d (%a)")
    md_table = tbl.to_markdown(index=False)

    extra = ""
    if stats["avg_water"] is not None:
        extra = f"  \\n- 平均每日飲水量：{stats['avg_water']:.1f} L"

    md = (
f"# 📊 減重週報（{week_tag}）\\n\\n"
f"**週期：{start_date.strftime('%Y/%m/%d')} ～ {end_date.strftime('%Y/%m/%d')}**  \\n\\n"
"---\\n\\n"
"## 📈 體重與體脂紀錄\\n\\n"
f"{md_table}\\n\\n"
"---\\n\\n"
"## 📊 趨勢圖\\n\\n"
f"![體重趨勢]({os.path.basename(png_weight)})\\n"
f"![體脂率趨勢]({os.path.basename(png_bodyfat)})\\n\\n"
"---\\n\\n"
"## 📌 本週統計\\n\\n"
f"- 體重（AM）：{stats['start_weight_am']:.1f} → {stats['end_weight_am']:.1f} kg  (**{stats['delta_weight_am']:+.1f} kg**), 週平均 {stats['avg_weight_am']:.1f} kg  \\n"
f"- 體重（PM）：{stats['start_weight_pm']:.1f} → {stats['end_weight_pm']:.1f} kg  (**{stats['delta_weight_pm']:+.1f} kg**), 週平均 {stats['avg_weight_pm']:.1f} kg  \\n"
f"- 體重（AM+PM 平均）：{stats['avg_weight_all']:.1f} kg  \\n\\n"
f"- 體脂（AM）：{stats['start_fat_am']:.1f}% → {stats['end_fat_am']:.1f}%  (**{stats['delta_fat_am']:+.1f}%**), 週平均 {stats['avg_fat_am']:.1f}%  \\n"
f"- 體脂（PM）：{stats['start_fat_pm']:.1f}% → {stats['end_fat_pm']:.1f}%  (**{stats['delta_fat_pm']:+.1f}%**), 週平均 {stats['avg_fat_pm']:.1f}%  \\n"
f"- 體脂（AM+PM 平均）：{stats['avg_fat_all']:.1f}%  \\n\\n"
f"- 紀錄天數：{stats['days']} 天{extra}\\n\\n"
"---\\n\\n"
"## ✅ 建議\\n"
"- 維持 **高蛋白 (每公斤 1.6–2.0 g)** 與 **每週 2–3 次阻力訓練**  \\n"
"- 飲水 **≥ 3 L/天**（依活動量調整）  \\n"
"- 若每週下降 > 2.5 kg，建議微調熱量或與醫師討論  \\n"
    )
    with open(out_md_path, "w", encoding="utf-8") as f:
        f.write(md)

def main():
    p = argparse.ArgumentParser(description="以週五為起始的自訂週期，從 master 產生 Excel + Markdown + 圖表（含中文字體修正）")
    p.add_argument("master", nargs="?", default="GLP1_weight_tracking_master.xlsx", help="主檔 Excel（手動維護）")
    p.add_argument("--sheet", default=None, help="工作表名稱（預設先嘗試 'Daily Log'，再退回第一個工作表）")
    p.add_argument("--header-row", type=int, default=0, help="欄位標題所在的列索引（0=第一列）")
    p.add_argument("--anchor-date", default="2025-08-15", help="每週起始的對齊基準日（週五），例如 2025-08-15")
    p.add_argument("--week-index", type=int, default=None, help="第幾週（以 anchor-date 為第1週起算）；未提供則取最後一週")
    p.add_argument("--out-root", default=".", help="輸出根目錄（會在裡面建立 weekly/ 與 reports/）")
    args = p.parse_args()

    df = read_daily_log(args.master, sheet_name=args.sheet, header_row=args.header_row)

    wdf, week_tag, start_date, end_date = pick_custom_week(df, args.anchor_date, args.week_index)

    weekly_dir = os.path.join(args.out_root, "weekly")
    reports_dir = os.path.join(args.out_root, "reports")
    ensure_dirs(weekly_dir); ensure_dirs(reports_dir)

    weekly_xlsx = os.path.join(weekly_dir, f"{week_tag}_weight_tracking.xlsx")
    save_weekly_excel(wdf, weekly_xlsx)

    weight_png, bodyfat_png = make_charts(wdf, reports_dir, prefix=week_tag)

    stats = compute_stats(wdf)
    weekly_md = os.path.join(reports_dir, f"{week_tag}_weekly_report.md")
    make_markdown(wdf, stats, weight_png, bodyfat_png, weekly_md, week_tag, start_date, end_date)

    print("✅ 已完成輸出")
    print("Weekly Excel:", weekly_xlsx)
    print("Report MD   :", weekly_md)
    print("Charts      :", weight_png, bodyfat_png)

if __name__ == "__main__":
    main()

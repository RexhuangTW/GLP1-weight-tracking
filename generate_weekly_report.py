
import argparse
import os
import re
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib

# -------- Chinese font fallback (for charts) --------
matplotlib.rcParams["font.sans-serif"] = [
    "Noto Sans CJK TC", "Noto Sans CJK SC", "Noto Sans CJK JP",
    "Microsoft JhengHei", "PingFang TC", "Heiti TC", "SimHei",
    "WenQuanYi Micro Hei", "Arial Unicode MS", "DejaVu Sans"
]
matplotlib.rcParams["axes.unicode_minus"] = False

# -------- Column definitions --------
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

# -------- Helpers --------
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
    """
    讀取數據源，支援 Excel 或 CSV 格式
    - 若為 CSV：自動將每日多次測量轉換為早上/晚上格式
    - 若為 Excel：使用原有的欄位映射邏輯
    """
    # 判斷文件類型
    if master_path.lower().endswith('.csv'):
        # 讀取 CSV 文件
        df_raw = pd.read_csv(master_path)
        
        # 解析測量日期時間
        df_raw['測量日期時間'] = pd.to_datetime(df_raw['測量日期'], format='%Y/%m/%d %H:%M')
        
        # 調整日期：凌晨 0:00-4:59 算作前一天
        df_raw['調整日期'] = df_raw['測量日期時間'].apply(
            lambda dt: (dt - pd.Timedelta(days=1)).date() if dt.hour < 5 else dt.date()
        )
        df_raw['日期'] = df_raw['調整日期']
        df_raw['時間'] = df_raw['測量日期時間'].dt.time
        df_raw['小時'] = df_raw['測量日期時間'].dt.hour
        
        # 分類早上/晚上：早上定義為 5:00-11:59，晚上為 12:00-4:59（隔天）
        df_raw['時段'] = df_raw['小時'].apply(lambda h: 'AM' if 5 <= h < 12 else 'PM')
        
        # 按日期和時段分組，取平均值（若一天有多次測量）
        daily_data = []
        for date in df_raw['日期'].unique():
            date_df = df_raw[df_raw['日期'] == date]
            
            am_data = date_df[date_df['時段'] == 'AM']
            pm_data = date_df[date_df['時段'] == 'PM']
            
            row = {'日期': pd.to_datetime(date)}
            
            # 早上數據
            if not am_data.empty:
                row['早上體重 (kg)'] = am_data['體重(kg)'].mean()
                row['早上體脂 (%)'] = am_data['體脂肪(%)'].mean()
                row['早上內臟脂肪'] = am_data['內臟脂肪程度'].mean()
            else:
                row['早上體重 (kg)'] = None
                row['早上體脂 (%)'] = None
                row['早上內臟脂肪'] = None
            
            # 晚上數據
            if not pm_data.empty:
                row['晚上體重 (kg)'] = pm_data['體重(kg)'].mean()
                row['晚上體脂 (%)'] = pm_data['體脂肪(%)'].mean()
                row['晚上內臟脂肪'] = pm_data['內臟脂肪程度'].mean()
            else:
                row['晚上體重 (kg)'] = None
                row['晚上體脂 (%)'] = None
                row['晚上內臟脂肪'] = None
            
            # 只添加至少有一個測量值的日期
            if row['早上體重 (kg)'] is not None or row['晚上體重 (kg)'] is not None:
                daily_data.append(row)
        
        df_final = pd.DataFrame(daily_data)
        df_final = df_final.sort_values('日期').reset_index(drop=True)
        
        # 添加可選欄位（CSV中沒有，設為None）
        for col in OPTIONAL_LOGICAL:
            if col not in df_final.columns:
                df_final[col] = None
        
        return df_final
    
    else:
        # Excel 格式：使用原有邏輯
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

def pandas_offset_weeks(n):
    return pd.Timedelta(days=7*n)

def assign_custom_week(df, anchor_date):
    d0 = pd.to_datetime(anchor_date).normalize()  # Friday anchor
    delta_days = (df["日期"].dt.normalize() - d0).dt.days
    week_idx = (delta_days // 7) + 1  # 1-based
    df2 = df.copy()
    df2["WEEK_IDX"] = week_idx
    return df2

def pick_custom_week(df, anchor_date, week_index=None):
    df2 = assign_custom_week(df, anchor_date)
    target = int(df2["WEEK_IDX"].max() if week_index is None else week_index)
    wdf = df2[df2["WEEK_IDX"] == target].copy()
    if wdf.empty:
        raise ValueError(f"在 anchor={anchor_date} 下，找不到第 {target} 週的資料。")
    start_date = pd.to_datetime(anchor_date) + pandas_offset_weeks(target-1)
    end_date = start_date + pd.Timedelta(days=6)
    tag = f"{start_date.year}-CW{target:02d}"
    return wdf, tag, start_date, end_date

def _first_last_valid(series):
    s = series.dropna()
    if s.empty:
        return None, None
    return float(s.iloc[0]), float(s.iloc[-1])

def _fmt(x, digits=1, unit=""):
    if x is None or (isinstance(x, float) and x != x):
        return "-"
    return f"{x:.{digits}f}" + (f" {unit}" if unit else "")

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

    # 內臟脂肪趨勢圖
    if '早上內臟脂肪' in wdf.columns and '晚上內臟脂肪' in wdf.columns:
        plt.figure(figsize=(8,5))
        plt.plot(wdf["日期"], wdf["早上內臟脂肪"], marker="o", label="早上內臟脂肪", color='#ff7f0e')
        plt.plot(wdf["日期"], wdf["晚上內臟脂肪"], marker="o", label="晚上內臟脂肪", color='#d62728')
        plt.xlabel("日期"); plt.ylabel("內臟脂肪程度"); plt.title("內臟脂肪趨勢"); plt.legend(); plt.grid(True)
        plt.xticks(rotation=30)
        # 添加健康參考線
        plt.axhline(y=10, color='green', linestyle='--', alpha=0.5, label='標準 (≤9.5)')
        plt.axhline(y=15, color='orange', linestyle='--', alpha=0.5, label='偏高 (10-14.5)')
        plt.legend()
        visceral_png = os.path.join(out_dir, f"{prefix}_visceral_fat_trend.png")
        plt.savefig(visceral_png, dpi=150, bbox_inches="tight"); plt.close()
    else:
        visceral_png = None

    return weight_png, bodyfat_png, visceral_png

def compute_stats(wdf):
    wdf_sorted = wdf.sort_values("日期")
    sw_am, ew_am = _first_last_valid(wdf_sorted["早上體重 (kg)"])
    sw_pm, ew_pm = _first_last_valid(wdf_sorted["晚上體重 (kg)"])
    sf_am, ef_am = _first_last_valid(wdf_sorted["早上體脂 (%)"])
    sf_pm, ef_pm = _first_last_valid(wdf_sorted["晚上體脂 (%)"])
    
    # 內臟脂肪統計
    sv_am, ev_am, sv_pm, ev_pm = None, None, None, None
    if '早上內臟脂肪' in wdf_sorted.columns and '晚上內臟脂肪' in wdf_sorted.columns:
        sv_am, ev_am = _first_last_valid(wdf_sorted["早上內臟脂肪"])
        sv_pm, ev_pm = _first_last_valid(wdf_sorted["晚上內臟脂肪"])

    stats = {
        "period_start": wdf_sorted["日期"].iloc[0].strftime("%Y/%m/%d"),
        "period_end":   wdf_sorted["日期"].iloc[-1].strftime("%Y/%m/%d"),
        "start_weight_am": sw_am,
        "end_weight_am":   ew_am,
        "delta_weight_am": (ew_am - sw_am) if (sw_am is not None and ew_am is not None) else None,
        "avg_weight_am":   float(wdf_sorted["早上體重 (kg)"].mean()),
        "start_weight_pm": sw_pm,
        "end_weight_pm":   ew_pm,
        "delta_weight_pm": (ew_pm - sw_pm) if (sw_pm is not None and ew_pm is not None) else None,
        "avg_weight_pm":   float(wdf_sorted["晚上體重 (kg)"].mean()),
        "avg_weight_all":  float(wdf_sorted[["早上體重 (kg)","晚上體重 (kg)"]].mean().mean()),
        "start_fat_am": sf_am,
        "end_fat_am":   ef_am,
        "delta_fat_am": (ef_am - sf_am) if (sf_am is not None and ef_am is not None) else None,
        "avg_fat_am":   float(wdf_sorted["早上體脂 (%)"].mean()),
        "start_fat_pm": sf_pm,
        "end_fat_pm":   ef_pm,
        "delta_fat_pm": (ef_pm - sf_pm) if (sf_pm is not None and ef_pm is not None) else None,
        "avg_fat_pm":   float(wdf_sorted["晚上體脂 (%)"].mean()),
        "avg_fat_all":  float(wdf_sorted[["早上體脂 (%)","晚上體脂 (%)"]].mean().mean()),
        "days": int(wdf_sorted.shape[0])
    }
    
    # 內臟脂肪統計
    if sv_am is not None or sv_pm is not None:
        stats["start_visceral_am"] = sv_am
        stats["end_visceral_am"] = ev_am
        stats["delta_visceral_am"] = (ev_am - sv_am) if (sv_am is not None and ev_am is not None) else None
        stats["avg_visceral_am"] = float(wdf_sorted["早上內臟脂肪"].mean())
        stats["start_visceral_pm"] = sv_pm
        stats["end_visceral_pm"] = ev_pm
        stats["delta_visceral_pm"] = (ev_pm - sv_pm) if (sv_pm is not None and ev_pm is not None) else None
        stats["avg_visceral_pm"] = float(wdf_sorted["晚上內臟脂肪"].mean())
        stats["avg_visceral_all"] = float(wdf_sorted[["早上內臟脂肪","晚上內臟脂肪"]].mean().mean())
    else:
        stats["start_visceral_am"] = None
        stats["end_visceral_am"] = None
        stats["delta_visceral_am"] = None
        stats["avg_visceral_am"] = None
        stats["start_visceral_pm"] = None
        stats["end_visceral_pm"] = None
        stats["delta_visceral_pm"] = None
        stats["avg_visceral_pm"] = None
        stats["avg_visceral_all"] = None
    
    if "每日飲水量 (L)" in wdf_sorted.columns:
        water = wdf_sorted["每日飲水量 (L)"].dropna()
        stats["avg_water"] = float(water.mean()) if not water.empty else None
    else:
        stats["avg_water"] = None
    return stats

def make_markdown(wdf, stats, png_weight, png_bodyfat, png_visceral, out_md_path, week_tag, start_date, end_date):
    # 基本表格
    table_cols = ["日期","早上體重 (kg)","晚上體重 (kg)","早上體脂 (%)","晚上體脂 (%)"]
    if '早上內臟脂肪' in wdf.columns and '晚上內臟脂肪' in wdf.columns:
        table_cols.extend(["早上內臟脂肪","晚上內臟脂肪"])
    
    tbl = wdf[table_cols].copy()

    weekday_zh = {0:"週一",1:"週二",2:"週三",3:"週四",4:"週五",5:"週六",6:"週日"}
    tbl["日期"] = tbl["日期"].apply(lambda d: d.strftime('%m/%d') + f" ({weekday_zh[d.weekday()]})")

    md_table = tbl.to_markdown(index=False)

    extra = ""
    if stats["avg_water"] is not None:
        extra = f"  \n- 平均每日飲水量：{_fmt(stats['avg_water'])} L"

    # 趨勢圖部分
    charts_section = (
        "## 📊 趨勢圖\n\n"
        f"![體重趨勢]({os.path.basename(png_weight)})\n"
        f"![體脂率趨勢]({os.path.basename(png_bodyfat)})\n"
    )
    if png_visceral:
        charts_section += f"![內臟脂肪趨勢]({os.path.basename(png_visceral)})\n"
    charts_section += "\n---\n\n"

    # 內臟脂肪統計
    visceral_stats = ""
    if stats.get("avg_visceral_am") is not None:
        visceral_stats = (
            f"\n- 內臟脂肪（AM）：{_fmt(stats['start_visceral_am'], 1)} → {_fmt(stats['end_visceral_am'], 1)}  "
            f"(**{_fmt(stats['delta_visceral_am'], 1)}**), 週平均 {stats['avg_visceral_am']:.1f}  \n"
            f"- 內臟脂肪（PM）：{_fmt(stats['start_visceral_pm'], 1)} → {_fmt(stats['end_visceral_pm'], 1)}  "
            f"(**{_fmt(stats['delta_visceral_pm'], 1)}**), 週平均 {stats['avg_visceral_pm']:.1f}  \n"
            f"- 內臟脂肪（AM+PM 平均）：{stats['avg_visceral_all']:.1f}  \n"
            f"  💡 *標準：≤9.5，偏高：10-14.5，過高：≥15*  \n"
        )

    md = (
        f"# 📊 減重週報（{week_tag}）\n\n"
        f"**週期：{start_date.strftime('%Y/%m/%d')} ～ {end_date.strftime('%Y/%m/%d')}**  \n\n"
        "---\n\n"
        "## 📈 體重與體脂紀錄\n\n"
        f"{md_table}\n\n"
        "---\n\n"
        f"{charts_section}"
        "## 📌 本週統計\n\n"
        f"- 體重（AM）：{_fmt(stats['start_weight_am'])} → {_fmt(stats['end_weight_am'])} kg  (**{_fmt(stats['delta_weight_am'])} kg**), 週平均 {stats['avg_weight_am']:.1f} kg  \n"
        f"- 體重（PM）：{_fmt(stats['start_weight_pm'])} → {_fmt(stats['end_weight_pm'])} kg  (**{_fmt(stats['delta_weight_pm'])} kg**), 週平均 {stats['avg_weight_pm']:.1f} kg  \n"
        f"- 體重（AM+PM 平均）：{stats['avg_weight_all']:.1f} kg  \n\n"
        f"- 體脂（AM）：{_fmt(stats['start_fat_am'])}% → {_fmt(stats['end_fat_am'])}%  (**{_fmt(stats['delta_fat_am'])}%**), 週平均 {stats['avg_fat_am']:.1f}%  \n"
        f"- 體脂（PM）：{_fmt(stats['start_fat_pm'])}% → {_fmt(stats['end_fat_pm'])}%  (**{_fmt(stats['delta_fat_pm'])}%**), 週平均 {stats['avg_fat_pm']:.1f}%  \n"
        f"- 體脂（AM+PM 平均）：{stats['avg_fat_all']:.1f}%  \n"
        f"{visceral_stats}\n"
        f"- 紀錄天數：{stats['days']} 天{extra}\n\n"
        "---\n\n"
        "## ✅ 建議\n"
        "- 維持 **高蛋白 (每公斤 1.6–2.0 g)** 與 **每週 2–3 次阻力訓練**  \n"
        "- 飲水 **≥ 3 L/天**（依活動量調整）  \n"
        "- 若每週下降 > 2.5 kg，建議微調熱量或與醫師討論  \n"
    )
    with open(out_md_path, "w", encoding="utf-8") as f:
        f.write(md)

def make_summary_report(df, out_dir, prefix="summary"):
    """產生從第一天到最新數據的總結報告"""
    df_sorted = df.sort_values("日期")
    
    # 計算整體統計
    stats = compute_stats(df_sorted)
    
    # 產生圖表
    weight_png, bodyfat_png, visceral_png = make_charts(df_sorted, out_dir, prefix=prefix)
    
    # 計算週次
    total_days = len(df_sorted)
    total_weeks = (total_days + 6) // 7  # 向上取整
    
    # 產生表格 - 只顯示最近7天和第一天作對比
    recent_data = df_sorted.tail(7)
    first_day = df_sorted.iloc[0:1].copy()
    
    # 表格欄位
    table_cols = ["日期","早上體重 (kg)","晚上體重 (kg)","早上體脂 (%)","晚上體脂 (%)"]
    has_visceral = '早上內臟脂肪' in df_sorted.columns and '晚上內臟脂肪' in df_sorted.columns
    if has_visceral:
        table_cols.extend(["早上內臟脂肪","晚上內臟脂肪"])
    
    if len(df_sorted) <= 7:
        display_data = df_sorted[table_cols].copy()
    else:
        # 創建分隔行
        separator_dict = {"日期": ["..."], "早上體重 (kg)": ["..."], "晚上體重 (kg)": ["..."], 
                         "早上體脂 (%)": ["..."], "晚上體脂 (%)": ["..."]}
        if has_visceral:
            separator_dict["早上內臟脂肪"] = ["..."]
            separator_dict["晚上內臟脂肪"] = ["..."]
        separator_row = pd.DataFrame(separator_dict)
        display_data = pd.concat([first_day[table_cols], separator_row, recent_data[table_cols]], ignore_index=True)
    
    # 格式化日期
    weekday_zh = {0:"週一",1:"週二",2:"週三",3:"週四",4:"週五",5:"週六",6:"週日"}
    display_data_copy = display_data.copy()
    
    for idx in display_data_copy.index:
        date_val = display_data_copy.loc[idx, "日期"]
        if date_val != "..." and pd.notna(date_val):
            display_data_copy.loc[idx, "日期"] = date_val.strftime('%m/%d') + f" ({weekday_zh[date_val.weekday()]})"
    
    md_table = display_data_copy.to_markdown(index=False)
    
    # 計算總體趨勢
    start_date = df_sorted["日期"].iloc[0]
    end_date = df_sorted["日期"].iloc[-1]
    
    # 額外統計
    extra = ""
    if stats["avg_water"] is not None:
        extra = f"  \n- 平均每日飲水量：{_fmt(stats['avg_water'])} L"
    
    # 週次分析
    weekly_analysis = ""
    if total_weeks > 1:
        weekly_weight_loss_am = stats['delta_weight_am'] / total_weeks if stats['delta_weight_am'] else 0
        weekly_weight_loss_pm = stats['delta_weight_pm'] / total_weeks if stats['delta_weight_pm'] else 0
        weekly_analysis = f"  \n- 平均每週體重變化（AM）：{_fmt(weekly_weight_loss_am)} kg/週  \n- 平均每週體重變化（PM）：{_fmt(weekly_weight_loss_pm)} kg/週"
    
    # 趨勢圖部分
    charts_section = (
        "## 📊 整體趨勢圖\n\n"
        f"![體重趨勢]({os.path.basename(weight_png)})\n"
        f"![體脂率趨勢]({os.path.basename(bodyfat_png)})\n"
    )
    if visceral_png:
        charts_section += f"![內臟脂肪趨勢]({os.path.basename(visceral_png)})\n"
    charts_section += "\n---\n\n"
    
    # 內臟脂肪統計
    visceral_stats = ""
    if stats.get("avg_visceral_am") is not None:
        visceral_stats = (
            f"\n- **內臟脂肪（AM）**：{_fmt(stats['start_visceral_am'], 1)} → {_fmt(stats['end_visceral_am'], 1)}  "
            f"(**{_fmt(stats['delta_visceral_am'], 1)}**), 總平均 {stats['avg_visceral_am']:.1f}  \n"
            f"- **內臟脂肪（PM）**：{_fmt(stats['start_visceral_pm'], 1)} → {_fmt(stats['end_visceral_pm'], 1)}  "
            f"(**{_fmt(stats['delta_visceral_pm'], 1)}**), 總平均 {stats['avg_visceral_pm']:.1f}  \n"
            f"- **內臟脂肪（AM+PM 平均）**：{stats['avg_visceral_all']:.1f}  \n"
            f"  💡 *標準：≤9.5，偏高：10-14.5，過高：≥15*  \n"
        )
    
    md = (
        f"# 📊 減重總結報告\n\n"
        f"**總期間：{start_date.strftime('%Y/%m/%d')} ～ {end_date.strftime('%Y/%m/%d')}**  \n"
        f"**追蹤期間：{total_days} 天 ({total_weeks} 週)**  \n\n"
        "---\n\n"
        "## 📈 體重與體脂紀錄概覽\n\n"
        "*顯示第一天與最近7天的數據*\n\n"
        f"{md_table}\n\n"
        "---\n\n"
        f"{charts_section}"
        "## 📌 總體統計\n\n"
        f"- **體重（AM）**：{_fmt(stats['start_weight_am'])} → {_fmt(stats['end_weight_am'])} kg  (**{_fmt(stats['delta_weight_am'])} kg**), 總平均 {stats['avg_weight_am']:.1f} kg  \n"
        f"- **體重（PM）**：{_fmt(stats['start_weight_pm'])} → {_fmt(stats['end_weight_pm'])} kg  (**{_fmt(stats['delta_weight_pm'])} kg**), 總平均 {stats['avg_weight_pm']:.1f} kg  \n"
        f"- **體重（AM+PM 平均）**：{stats['avg_weight_all']:.1f} kg  \n\n"
        f"- **體脂（AM）**：{_fmt(stats['start_fat_am'])}% → {_fmt(stats['end_fat_am'])}%  (**{_fmt(stats['delta_fat_am'])}%**), 總平均 {stats['avg_fat_am']:.1f}%  \n"
        f"- **體脂（PM）**：{_fmt(stats['start_fat_pm'])}% → {_fmt(stats['end_fat_pm'])}%  (**{_fmt(stats['delta_fat_pm'])}%**), 總平均 {stats['avg_fat_pm']:.1f}%  \n"
        f"- **體脂（AM+PM 平均）**：{stats['avg_fat_all']:.1f}%  \n"
        f"{visceral_stats}\n"
        f"- **追蹤天數**：{stats['days']} 天{extra}{weekly_analysis}\n\n"
        "---\n\n"
        "## 🎯 重點成果\n\n"
    )
    
    # 成果分析
    if stats['delta_weight_am'] and stats['delta_weight_am'] < 0:
        md += f"✅ **體重減少**：在 {total_days} 天內減重 {abs(stats['delta_weight_am']):.1f} kg（早上測量）  \n"
    if stats['delta_fat_pm'] and stats['delta_fat_pm'] < 0:
        md += f"✅ **體脂下降**：體脂率降低 {abs(stats['delta_fat_pm']):.1f}%（晚上測量）  \n"
    if stats.get('delta_visceral_am') and stats['delta_visceral_am'] < 0:
        md += f"✅ **內臟脂肪改善**：內臟脂肪程度降低 {abs(stats['delta_visceral_am']):.1f}（早上測量）  \n"
    
    md += "\n## ✅ 持續建議\n"
    md += "- 維持 **高蛋白 (每公斤 1.6–2.0 g)** 與 **每週 2–3 次阻力訓練**  \n"
    md += "- 飲水 **≥ 3 L/天**（依活動量調整）  \n"
    md += "- 持續監測體重與體脂變化，建議保持每週穩定減重  \n"
    md += "- 如有任何異常變化，建議諮詢專業醫師  \n"
    
    return md, weight_png, bodyfat_png, visceral_png

def main():
    p = argparse.ArgumentParser(description="以週五為起始的自訂週期，從 master 產生 Excel + Markdown + 圖表（支援 CSV/Excel 格式）")
    p.add_argument("master", nargs="?", default="BodyComposition_202507-202510.csv", help="主檔（CSV 或 Excel 格式）")
    p.add_argument("--sheet", default=None, help="工作表名稱（僅用於 Excel，預設先嘗試 'Daily Log'，再退回第一個工作表）")
    p.add_argument("--header-row", type=int, default=0, help="欄位標題所在的列索引（僅用於 Excel，0=第一列）")
    p.add_argument("--anchor-date", default="2025-08-15", help="每週起始的對齊基準日（週四），例如 2025-08-15")
    p.add_argument("--week-index", type=int, default=None, help="第幾週（以 anchor-date 為第1週起算）；未提供則取最後一週")
    p.add_argument("--out-root", default=".", help="輸出根目錄（會在裡面建立 weekly/ 與 reports/）")
    p.add_argument("--summary", action="store_true", help="產生從第一天到最新數據的總結報告")
    args = p.parse_args()

    df = read_daily_log(args.master, sheet_name=args.sheet, header_row=args.header_row)

    if args.summary:
        # 產生總結報告
        reports_dir = os.path.join(args.out_root, "reports")
        summary_dir = os.path.join(reports_dir, "summary")
        ensure_dirs(summary_dir)
        
        summary_md, weight_png, bodyfat_png, visceral_png = make_summary_report(df, summary_dir)
        summary_md_path = os.path.join(summary_dir, "overall_summary_report.md")
        
        with open(summary_md_path, "w", encoding="utf-8") as f:
            f.write(summary_md)
        
        print("✅ 總結報告已完成輸出")
        print("Summary MD :", summary_md_path)
        charts_list = [weight_png, bodyfat_png]
        if visceral_png:
            charts_list.append(visceral_png)
        print("Charts     :", " ".join(charts_list))
        return

    wdf, week_tag, start_date, end_date = pick_custom_week(df, args.anchor_date, args.week_index)

    weekly_dir = os.path.join(args.out_root, "weekly")
    reports_dir = os.path.join(args.out_root, "reports")
    week_reports_dir = os.path.join(reports_dir, week_tag)  # 在 reports 下建立週期子資料夾
    ensure_dirs(weekly_dir); ensure_dirs(week_reports_dir)

    weekly_xlsx = os.path.join(weekly_dir, f"{week_tag}_weight_tracking.xlsx")
    save_weekly_excel(wdf, weekly_xlsx)

    weight_png, bodyfat_png, visceral_png = make_charts(wdf, week_reports_dir, prefix=week_tag)

    stats = compute_stats(wdf)
    weekly_md = os.path.join(week_reports_dir, f"{week_tag}_weekly_report.md")
    make_markdown(wdf, stats, weight_png, bodyfat_png, visceral_png, weekly_md, week_tag, start_date, end_date)

    print("✅ 已完成輸出")
    print("Weekly Excel:", weekly_xlsx)
    print("Report MD   :", weekly_md)
    charts_list = [weight_png, bodyfat_png]
    if visceral_png:
        charts_list.append(visceral_png)
    print("Charts      :", " ".join(charts_list))

if __name__ == "__main__":
    main()


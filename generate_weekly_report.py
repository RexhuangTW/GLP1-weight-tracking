
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
                row['早上骨骼肌 (%)'] = am_data['骨骼肌(%)'].mean()
                # 計算脂肪重量和骨骼肌重量
                row['早上脂肪重量 (kg)'] = row['早上體重 (kg)'] * row['早上體脂 (%)'] / 100
                row['早上骨骼肌重量 (kg)'] = row['早上體重 (kg)'] * row['早上骨骼肌 (%)'] / 100
            else:
                row['早上體重 (kg)'] = None
                row['早上體脂 (%)'] = None
                row['早上內臟脂肪'] = None
                row['早上骨骼肌 (%)'] = None
                row['早上脂肪重量 (kg)'] = None
                row['早上骨骼肌重量 (kg)'] = None
            
            # 晚上數據
            if not pm_data.empty:
                row['晚上體重 (kg)'] = pm_data['體重(kg)'].mean()
                row['晚上體脂 (%)'] = pm_data['體脂肪(%)'].mean()
                row['晚上內臟脂肪'] = pm_data['內臟脂肪程度'].mean()
                row['晚上骨骼肌 (%)'] = pm_data['骨骼肌(%)'].mean()
                # 計算脂肪重量和骨骼肌重量
                row['晚上脂肪重量 (kg)'] = row['晚上體重 (kg)'] * row['晚上體脂 (%)'] / 100
                row['晚上骨骼肌重量 (kg)'] = row['晚上體重 (kg)'] * row['晚上骨骼肌 (%)'] / 100
            else:
                row['晚上體重 (kg)'] = None
                row['晚上體脂 (%)'] = None
                row['晚上內臟脂肪'] = None
                row['晚上骨骼肌 (%)'] = None
                row['晚上脂肪重量 (kg)'] = None
                row['晚上骨骼肌重量 (kg)'] = None
            
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

def pick_month(df: pd.DataFrame, month_str: str | None):
    """選取某個月份的資料，回傳 (wdf, ym_tag, start_date, end_date)
    - month_str 形如 'YYYY-MM'；若為 None，則取 df 中最新日期所屬月份
    """
    sdf = df.copy()
    # 確保是 datetime（方便取年月）
    sdf['日期_dt'] = pd.to_datetime(sdf['日期'])
    if month_str is None:
        latest = sdf['日期_dt'].max()
        ym = latest.strftime('%Y-%m')
    else:
        ym = month_str
    year, month = map(int, ym.split('-'))
    mask = (sdf['日期_dt'].dt.year == year) & (sdf['日期_dt'].dt.month == month)
    wdf = sdf.loc[mask].copy()
    if wdf.empty:
        raise ValueError(f"指定月份 {ym} 沒有資料")
    wdf = wdf.drop(columns=['日期_dt'])
    start_date = pd.to_datetime(wdf['日期']).min().date()
    end_date = pd.to_datetime(wdf['日期']).max().date()
    ym_tag = f"{year}-{month:02d}"
    return wdf, ym_tag, start_date, end_date

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

def make_charts(wdf, out_dir, prefix, kpi=None, is_week=False, show_ma: bool = False, show_targets: bool = True):
    wdf_sorted = wdf.sort_values("日期")
    plt.figure(figsize=(8,5))
    plt.plot(wdf_sorted["日期"], wdf_sorted["早上體重 (kg)"], marker="o", label="早上體重")
    plt.plot(wdf_sorted["日期"], wdf_sorted["晚上體重 (kg)"], marker="o", label="晚上體重")
    # 7日移動平均（AM）
    if show_ma:
        if "早上體重 (kg)" in wdf_sorted.columns:
            ma = wdf_sorted["早上體重 (kg)"].rolling(window=7, min_periods=3).mean()
            plt.plot(wdf_sorted["日期"], ma, color="#1f77b4", linestyle=":", linewidth=2, alpha=0.9, label="7日均線(AM)")
    # 目標線：每週體重下降目標（線性）
    if show_targets and is_week and kpi and kpi.get("weight_target_end") is not None and kpi.get("weight_start") is not None:
        dates = list(wdf_sorted["日期"]) 
        n = len(dates)
        if n >= 2:
            y0 = kpi["weight_start"]
            y1 = kpi["weight_target_end"]
            y_line = [y0 + (y1 - y0) * i / (n - 1) for i in range(n)]
            plt.plot(dates, y_line, linestyle='--', color='#444', alpha=0.7, label=f"目標體重線 ({y1:.1f} kg)")
    plt.xlabel("日期"); plt.ylabel("體重 (kg)"); plt.title("體重趨勢"); plt.legend(); plt.grid(True)
    plt.xticks(rotation=30)
    weight_png = os.path.join(out_dir, f"{prefix}_weight_trend.png")
    plt.savefig(weight_png, dpi=150, bbox_inches="tight"); plt.close()

    plt.figure(figsize=(8,5))
    plt.plot(wdf_sorted["日期"], wdf_sorted["早上體脂 (%)"], marker="o", label="早上體脂")
    plt.plot(wdf_sorted["日期"], wdf_sorted["晚上體脂 (%)"], marker="o", label="晚上體脂")
    if show_ma and "早上體脂 (%)" in wdf_sorted.columns:
        ma = wdf_sorted["早上體脂 (%)"].rolling(window=7, min_periods=3).mean()
        plt.plot(wdf_sorted["日期"], ma, color="#ff1493", linestyle=":", linewidth=2, alpha=0.9, label="7日均線(AM)")
    # 目標線：體脂率每週下降目標（線性，以 AM 為主）
    if show_targets and is_week and kpi and kpi.get("fat_pct_target_end") is not None and kpi.get("fat_pct_start") is not None:
        dates = list(wdf_sorted["日期"]) 
        n = len(dates)
        if n >= 2:
            y0 = kpi["fat_pct_start"]
            y1 = kpi["fat_pct_target_end"]
            y_line = [y0 + (y1 - y0) * i / (n - 1) for i in range(n)]
            plt.plot(dates, y_line, linestyle='--', color='#888', alpha=0.7, label=f"目標體脂線 ({y1:.1f}%)")
    plt.xlabel("日期"); plt.ylabel("體脂 (%)"); plt.title("體脂趨勢"); plt.legend(); plt.grid(True)
    plt.xticks(rotation=30)
    bodyfat_png = os.path.join(out_dir, f"{prefix}_bodyfat_trend.png")
    plt.savefig(bodyfat_png, dpi=150, bbox_inches="tight"); plt.close()

    # 內臟脂肪趨勢圖
    if '早上內臟脂肪' in wdf.columns and '晚上內臟脂肪' in wdf.columns:
        plt.figure(figsize=(8,5))
        plt.plot(wdf_sorted["日期"], wdf_sorted["早上內臟脂肪"], marker="o", label="早上內臟脂肪", color='#ff7f0e')
        plt.plot(wdf_sorted["日期"], wdf_sorted["晚上內臟脂肪"], marker="o", label="晚上內臟脂肪", color='#d62728')
        if show_ma and "早上內臟脂肪" in wdf_sorted.columns:
            ma = wdf_sorted["早上內臟脂肪"].rolling(window=7, min_periods=3).mean()
            plt.plot(wdf_sorted["日期"], ma, color="#ff7f0e", linestyle=":", linewidth=2, alpha=0.8, label="7日均線(AM)")
        plt.xlabel("日期"); plt.ylabel("內臟脂肪程度"); plt.title("內臟脂肪趨勢"); plt.legend(); plt.grid(True)
        plt.xticks(rotation=30)
        # 添加健康參考線
        plt.axhline(y=10, color='green', linestyle='--', alpha=0.5, label='標準 (≤9.5)')
        plt.axhline(y=15, color='orange', linestyle='--', alpha=0.5, label='偏高 (10-14.5)')
        # 目標線：內臟脂肪每週下降目標（線性，以 AM 為主）
        if show_targets and is_week and kpi and kpi.get("visceral_target_end") is not None and kpi.get("visceral_start") is not None:
            dates = list(wdf_sorted["日期"]) 
            n = len(dates)
            if n >= 2:
                y0 = kpi["visceral_start"]
                y1 = kpi["visceral_target_end"]
                y_line = [y0 + (y1 - y0) * i / (n - 1) for i in range(n)]
                plt.plot(dates, y_line, linestyle='--', color='#4444aa', alpha=0.7, label=f"目標內臟脂肪線 ({y1:.1f})")
        plt.legend()
        visceral_png = os.path.join(out_dir, f"{prefix}_visceral_fat_trend.png")
        plt.savefig(visceral_png, dpi=150, bbox_inches="tight"); plt.close()
    else:
        visceral_png = None

    # 骨骼肌趨勢圖
    if '早上骨骼肌 (%)' in wdf.columns and '晚上骨骼肌 (%)' in wdf.columns:
        plt.figure(figsize=(8,5))
        plt.plot(wdf_sorted["日期"], wdf_sorted["早上骨骼肌 (%)"], marker="o", label="早上骨骼肌", color='#2ca02c')
        plt.plot(wdf_sorted["日期"], wdf_sorted["晚上骨骼肌 (%)"], marker="o", label="晚上骨骼肌", color='#8c564b')
        if show_ma and "早上骨骼肌 (%)" in wdf_sorted.columns:
            ma = wdf_sorted["早上骨骼肌 (%)"].rolling(window=7, min_periods=3).mean()
            plt.plot(wdf_sorted["日期"], ma, color="#2ca02c", linestyle=":", linewidth=2, alpha=0.8, label="7日均線(AM)")
        plt.xlabel("日期"); plt.ylabel("骨骼肌 (%)"); plt.title("骨骼肌趨勢"); plt.legend(); plt.grid(True)
        plt.xticks(rotation=30)
        # 目標線：骨骼肌率維持或微升（以 AM 為主，畫水平線）
        if show_targets and is_week and kpi and kpi.get("muscle_pct_floor") is not None:
            plt.axhline(y=kpi["muscle_pct_floor"], color='#2ca02c', linestyle='--', alpha=0.5, label=f"骨骼肌最低目標 ({kpi['muscle_pct_floor']:.1f}%)")
        muscle_png = os.path.join(out_dir, f"{prefix}_muscle_trend.png")
        plt.savefig(muscle_png, dpi=150, bbox_inches="tight"); plt.close()
    else:
        muscle_png = None

    return weight_png, bodyfat_png, visceral_png, muscle_png

# ---- Composition quality helper ----
def compute_quality_ratio(wdf, days: int = 28):
    """Compute recent fat-loss to weight-loss ratio over the last N days using AM values.
    Returns (ratio, details_dict) where ratio can be None if insufficient data.
    """
    if "日期" not in wdf.columns:
        return None, {}
    df = wdf.sort_values("日期").copy()
    if df.empty:
        return None, {}
    try:
        last_day = df["日期"].iloc[-1]
    except Exception:
        return None, {}
    import datetime as _dt
    start_cut = last_day - _dt.timedelta(days=days - 1)
    win = df[df["日期"] >= start_cut]
    def _first_last(series):
        s = series.dropna()
        if s.empty:
            return None, None
        return s.iloc[0], s.iloc[-1]
    sw, ew = _first_last(win.get("早上體重 (kg)", win.get("晚上體重 (kg)")))
    sfw, efw = _first_last(win.get("早上脂肪重量 (kg)", win.get("晚上脂肪重量 (kg)")))
    if sw is None or ew is None or sfw is None or efw is None:
        return None, {"window_days": days, "count": len(win)}
    weight_drop = max(0.0, sw - ew)
    fat_drop = max(0.0, sfw - efw)
    ratio = (fat_drop / weight_drop) if weight_drop > 1e-6 else None
    return ratio, {
        "window_days": days,
        "count": int(len(win)),
        "start_weight": sw,
        "end_weight": ew,
        "start_fat_weight": sfw,
        "end_fat_weight": efw,
        "weight_drop": weight_drop,
        "fat_drop": fat_drop,
    }

def _compute_eta_to_target(wdf, col_am: str, col_pm: str, target: float, days: int = 28, direction: str = 'down'):
    """Estimate days to reach target using last N days trend (AM preferred, fallback to PM).
    direction: 'down' for decreasing targets, 'up' for increasing.
    Returns dict or None: { 'days': int, 'weeks': float, 'date': datetime.date }
    """
    if target is None:
        return None
    df = wdf.sort_values("日期").copy()
    if df.empty or col_am not in df.columns:
        return None
    import datetime as _dt
    # choose series (AM preferred, fallback PM)
    series = df[col_am]
    if series.dropna().empty and col_pm in df.columns:
        series = df[col_pm]
    # window slice
    last_date = df["日期"].iloc[-1]
    start_cut = last_date - _dt.timedelta(days=days - 1)
    win = df[df["日期"] >= start_cut]
    y = series.loc[win.index].dropna()
    if y.empty:
        return None
    # find first and last valid within window and their dates
    first_idx = y.index[0]
    last_idx = y.index[-1]
    y0 = float(y.iloc[0]); y1 = float(y.iloc[-1])
    t0 = df.loc[first_idx, "日期"]; t1 = df.loc[last_idx, "日期"]
    dt_days = max(1.0, (t1 - t0).days)
    slope_per_day = (y1 - y0) / dt_days
    # direction check
    if direction == 'down':
        # need negative slope and target below current
        if not (slope_per_day < 0 and y1 > target):
            return None
        days_needed = (target - y1) / slope_per_day  # both negative -> positive
    else:
        if not (slope_per_day > 0 and y1 < target):
            return None
        days_needed = (target - y1) / slope_per_day
    if days_needed is None or days_needed <= 0:
        return None
    eta_days = int(round(days_needed))
    eta_date = t1 + _dt.timedelta(days=eta_days)
    return {"days": eta_days, "weeks": eta_days / 7.0, "date": eta_date.date()}

def _compute_eta(wdf_all, wdf_slice, metric: str, target: float, scope: str = 'global', method: str = 'regress28'):
    """Dispatch ETA computation by metric and scope. Uses 28-day linear regression on AM series (fallback PM).
    metric: 'fatkg' | 'weight' | 'fatpct'
    scope: 'global' -> window based on wdf_all; 'local' -> window based on wdf_slice
    """
    import numpy as np
    import datetime as _dt
    if target is None:
        return None
    dfbase = wdf_all if scope == 'global' else wdf_slice
    if dfbase is None or dfbase.empty:
        return None
    df = dfbase.sort_values('日期').copy()
    last_date = df['日期'].iloc[-1]
    if method in ('regress28','endpoint28'):
        start_cut = last_date - _dt.timedelta(days=27)
        win = df[df['日期'] >= start_cut]
    else:
        # all history window
        win = df
    # choose columns
    if metric == 'fatkg':
        col_am, col_pm = '早上脂肪重量 (kg)', '晚上脂肪重量 (kg)'
        direction = 'down'
    elif metric == 'weight':
        col_am, col_pm = '早上體重 (kg)', '晚上體重 (kg)'
        direction = 'down'
    else:
        col_am, col_pm = '早上體脂 (%)', '晚上體脂 (%)'
        direction = 'down'
    # series pick (AM preferred, drop NaNs; fallback to PM if AM無有效值)
    y = win[col_am] if col_am in win.columns else None
    if y is not None:
        y = y.dropna()
    if y is None or y.empty:
        y = win[col_pm] if col_pm in win.columns else None
        if y is not None:
            y = y.dropna()
    if y is None or y.empty:
        return None
    # 將日期與有效值對齊
    xdates = win['日期'].loc[y.index]
    if xdates.empty:
        return None
    # convert dates to day offsets
    x0 = xdates.iloc[0]
    x = (xdates - x0).dt.days.to_numpy()
    yy = y.to_numpy(dtype=float)
    # 至少需要兩個不同時間點
    if len(x) < 2 or (x[-1] - x[0]) == 0:
        return None
    if method.startswith('endpoint') or len(x) < 3:
        # endpoint slope
        a = (yy[-1] - yy[0]) / max(1.0, float(x[-1] - x[0]))
        b = yy[0] - a * x[0]
    else:
        # linear regression: y = a*x + b
        A = np.vstack([x, np.ones_like(x)]).T
        a, b = np.linalg.lstsq(A, yy, rcond=None)[0]
    # current value at last point
    cur = float(yy[-1])
    if direction == 'down' and not (a < 0 and cur > target):
        return None
    if direction == 'up' and not (a > 0 and cur < target):
        return None
    # solve a*x + b = target -> x = (target - b)/a
    x_target = (target - b) / a
    days_needed = x_target - x[-1]
    if not np.isfinite(days_needed) or days_needed <= 0:
        return None
    eta_days = int(round(days_needed))
    eta_date = last_date + _dt.timedelta(days=eta_days)
    return {"days": eta_days, "weeks": eta_days / 7.0, "date": eta_date.date()}

# ---- KPI helpers ----
def compute_weekly_kpi(stats: dict) -> dict:
    """Compute default weekly KPI targets based on start values.
    Targets are intentionally conservative to preserve muscle.
    """
    kpi = {}
    # weight: target weekly drop ~0.8 kg
    ws = stats.get('start_weight_am')
    if ws is None:
        ws = stats.get('start_weight_pm')
    if ws is not None:
        kpi['weight_start'] = ws
        kpi['weight_target_end'] = ws - 0.8
    # fat percent (AM): weekly drop target ~0.4 pp
    fps = stats.get('start_fat_am')
    if fps is None:
        fps = stats.get('start_fat_pm')
    if fps is not None:
        kpi['fat_pct_start'] = fps
        kpi['fat_pct_target_end'] = max(fps - 0.4, 0)
    # visceral: weekly drop target ~0.5 level
    vs = stats.get('start_visceral_am')
    if vs is None:
        vs = stats.get('start_visceral_pm')
    if vs is not None:
        kpi['visceral_start'] = vs
        kpi['visceral_target_end'] = max(vs - 0.5, 0)
    # muscle percent floor: keep at least start (AM)
    mps = stats.get('start_muscle_am')
    if mps is None:
        mps = stats.get('start_muscle_pm')
    if mps is not None:
        kpi['muscle_pct_floor'] = mps
    # muscle weight target: >= 0 change (display only)
    mw = stats.get('start_muscle_weight_am')
    if mw is None:
        mw = stats.get('start_muscle_weight_pm')
    if mw is not None:
        kpi['muscle_weight_start'] = mw
        kpi['muscle_weight_target_end'] = mw  # hold
    return kpi

def _progress_bar(current: float, target_delta: float, achieved_delta: float, width: int = 20, inverse: bool = False) -> str:
    """Render a simple textual progress bar for Markdown.
    If inverse is True, lower is better (e.g., weight/fat drops), so we use achieved/target.
    target_delta should be positive magnitude (e.g., 0.8 kg drop -> 0.8).
    achieved_delta is positive magnitude of improvement.
    """
    if target_delta is None or target_delta <= 0:
        return "(無目標)"
    ratio = max(0.0, min(1.0, achieved_delta / target_delta))
    filled = int(round(width * ratio))
    bar = '█' * filled + '░' * (width - filled)
    return f"[{bar}] {ratio*100:.0f}%"

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
    
    # 骨骼肌統計
    sm_am, em_am, sm_pm, em_pm = None, None, None, None
    if '早上骨骼肌 (%)' in wdf_sorted.columns and '晚上骨骼肌 (%)' in wdf_sorted.columns:
        sm_am, em_am = _first_last_valid(wdf_sorted["早上骨骼肌 (%)"])
        sm_pm, em_pm = _first_last_valid(wdf_sorted["晚上骨骼肌 (%)"])
        
        stats["start_muscle_am"] = sm_am
        stats["end_muscle_am"] = em_am
        stats["delta_muscle_am"] = (em_am - sm_am) if (sm_am is not None and em_am is not None) else None
        stats["avg_muscle_am"] = float(wdf_sorted["早上骨骼肌 (%)"].mean())
        stats["start_muscle_pm"] = sm_pm
        stats["end_muscle_pm"] = em_pm
        stats["delta_muscle_pm"] = (em_pm - sm_pm) if (sm_pm is not None and em_pm is not None) else None
        stats["avg_muscle_pm"] = float(wdf_sorted["晚上骨骼肌 (%)"].mean())
        stats["avg_muscle_all"] = float(wdf_sorted[["早上骨骼肌 (%)","晚上骨骼肌 (%)"]].mean().mean())
    else:
        stats["start_muscle_am"] = None
        stats["end_muscle_am"] = None
        stats["delta_muscle_am"] = None
        stats["avg_muscle_am"] = None
        stats["start_muscle_pm"] = None
        stats["end_muscle_pm"] = None
        stats["delta_muscle_pm"] = None
        stats["avg_muscle_pm"] = None
        stats["avg_muscle_all"] = None
    
    # 脂肪重量統計
    if "早上脂肪重量 (kg)" in wdf_sorted.columns and "晚上脂肪重量 (kg)" in wdf_sorted.columns:
        sfw_am, efw_am = _first_last_valid(wdf_sorted["早上脂肪重量 (kg)"])
        sfw_pm, efw_pm = _first_last_valid(wdf_sorted["晚上脂肪重量 (kg)"])
        stats["start_fat_weight_am"] = sfw_am
        stats["end_fat_weight_am"] = efw_am
        stats["delta_fat_weight_am"] = (efw_am - sfw_am) if (sfw_am is not None and efw_am is not None) else None
        stats["avg_fat_weight_am"] = float(wdf_sorted["早上脂肪重量 (kg)"].mean())
        stats["start_fat_weight_pm"] = sfw_pm
        stats["end_fat_weight_pm"] = efw_pm
        stats["delta_fat_weight_pm"] = (efw_pm - sfw_pm) if (sfw_pm is not None and efw_pm is not None) else None
        stats["avg_fat_weight_pm"] = float(wdf_sorted["晚上脂肪重量 (kg)"].mean())
        stats["avg_fat_weight_all"] = float(wdf_sorted[["早上脂肪重量 (kg)","晚上脂肪重量 (kg)"]].mean().mean())
    else:
        stats["start_fat_weight_am"] = None
        stats["end_fat_weight_am"] = None
        stats["delta_fat_weight_am"] = None
        stats["avg_fat_weight_am"] = None
        stats["start_fat_weight_pm"] = None
        stats["end_fat_weight_pm"] = None
        stats["delta_fat_weight_pm"] = None
        stats["avg_fat_weight_pm"] = None
        stats["avg_fat_weight_all"] = None
    
    # 骨骼肌重量統計
    if "早上骨骼肌重量 (kg)" in wdf_sorted.columns and "晚上骨骼肌重量 (kg)" in wdf_sorted.columns:
        smw_am, emw_am = _first_last_valid(wdf_sorted["早上骨骼肌重量 (kg)"])
        smw_pm, emw_pm = _first_last_valid(wdf_sorted["晚上骨骼肌重量 (kg)"])
        stats["start_muscle_weight_am"] = smw_am
        stats["end_muscle_weight_am"] = emw_am
        stats["delta_muscle_weight_am"] = (emw_am - smw_am) if (smw_am is not None and emw_am is not None) else None
        stats["avg_muscle_weight_am"] = float(wdf_sorted["早上骨骼肌重量 (kg)"].mean())
        stats["start_muscle_weight_pm"] = smw_pm
        stats["end_muscle_weight_pm"] = emw_pm
        stats["delta_muscle_weight_pm"] = (emw_pm - smw_pm) if (smw_pm is not None and emw_pm is not None) else None
        stats["avg_muscle_weight_pm"] = float(wdf_sorted["晚上骨骼肌重量 (kg)"].mean())
        stats["avg_muscle_weight_all"] = float(wdf_sorted[["早上骨骼肌重量 (kg)","晚上骨骼肌重量 (kg)"]].mean().mean())
    else:
        stats["start_muscle_weight_am"] = None
        stats["end_muscle_weight_am"] = None
        stats["delta_muscle_weight_am"] = None
        stats["avg_muscle_weight_am"] = None
        stats["start_muscle_weight_pm"] = None
        stats["end_muscle_weight_pm"] = None
        stats["delta_muscle_weight_pm"] = None
        stats["avg_muscle_weight_pm"] = None
        stats["avg_muscle_weight_all"] = None
    
    if "每日飲水量 (L)" in wdf_sorted.columns:
        water = wdf_sorted["每日飲水量 (L)"].dropna()
        stats["avg_water"] = float(water.mean()) if not water.empty else None
    else:
        stats["avg_water"] = None
    return stats

def make_markdown(wdf, stats, png_weight, png_bodyfat, png_visceral, png_muscle, out_md_path, week_tag, start_date, end_date, kpi_period_label="本週", goals: dict | None = None, eta_config: dict | None = None):
    # 基本表格
    table_cols = ["日期","早上體重 (kg)","晚上體重 (kg)","早上體脂 (%)","晚上體脂 (%)"]
    if '早上內臟脂肪' in wdf.columns and '晚上內臟脂肪' in wdf.columns:
        table_cols.extend(["早上內臟脂肪","晚上內臟脂肪"])
    if '早上骨骼肌 (%)' in wdf.columns and '晚上骨骼肌 (%)' in wdf.columns:
        table_cols.extend(["早上骨骼肌 (%)","晚上骨骼肌 (%)"])
    
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
    if png_muscle:
        charts_section += f"![骨骼肌趨勢]({os.path.basename(png_muscle)})\n"
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
    
    # 骨骼肌統計
    muscle_stats = ""
    if stats.get("avg_muscle_am") is not None:
        muscle_stats = (
            f"\n- 骨骼肌（AM）：{_fmt(stats['start_muscle_am'], 1)}% → {_fmt(stats['end_muscle_am'], 1)}%  "
            f"(**{_fmt(stats['delta_muscle_am'], 1)}%**), 週平均 {stats['avg_muscle_am']:.1f}%  \n"
            f"- 骨骼肌（PM）：{_fmt(stats['start_muscle_pm'], 1)}% → {_fmt(stats['end_muscle_pm'], 1)}%  "
            f"(**{_fmt(stats['delta_muscle_pm'], 1)}%**), 週平均 {stats['avg_muscle_pm']:.1f}%  \n"
            f"- 骨骼肌（AM+PM 平均）：{stats['avg_muscle_all']:.1f}%  \n"
        )
    
    # 脂肪重量統計
    fat_weight_stats = ""
    if stats.get("avg_fat_weight_am") is not None:
        fat_weight_stats = (
            f"\n- 脂肪重量（AM）：{_fmt(stats['start_fat_weight_am'], 1)} → {_fmt(stats['end_fat_weight_am'], 1)} kg  "
            f"(**{_fmt(stats['delta_fat_weight_am'], 1)} kg**), 週平均 {stats['avg_fat_weight_am']:.1f} kg  \n"
            f"- 脂肪重量（PM）：{_fmt(stats['start_fat_weight_pm'], 1)} → {_fmt(stats['end_fat_weight_pm'], 1)} kg  "
            f"(**{_fmt(stats['delta_fat_weight_pm'], 1)} kg**), 週平均 {stats['avg_fat_weight_pm']:.1f} kg  \n"
            f"- 脂肪重量（AM+PM 平均）：{stats['avg_fat_weight_all']:.1f} kg  \n"
        )
    
    # 骨骼肌重量統計
    muscle_weight_stats = ""
    if stats.get("avg_muscle_weight_am") is not None:
        muscle_weight_stats = (
            f"\n- 骨骼肌重量（AM）：{_fmt(stats['start_muscle_weight_am'], 1)} → {_fmt(stats['end_muscle_weight_am'], 1)} kg  "
            f"(**{_fmt(stats['delta_muscle_weight_am'], 1)} kg**), 週平均 {stats['avg_muscle_weight_am']:.1f} kg  \n"
            f"- 骨骼肌重量（PM）：{_fmt(stats['start_muscle_weight_pm'], 1)} → {_fmt(stats['end_muscle_weight_pm'], 1)} kg  "
            f"(**{_fmt(stats['delta_muscle_weight_pm'], 1)} kg**), 週平均 {stats['avg_muscle_weight_pm']:.1f} kg  \n"
            f"- 骨骼肌重量（AM+PM 平均）：{stats['avg_muscle_weight_all']:.1f} kg  \n"
        )

    md = (
        f"# 📊 減重週報（{week_tag})\n\n"
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
        f"{visceral_stats}"
        f"{muscle_stats}"
        f"{fat_weight_stats}"
        f"{muscle_weight_stats}\n"
        f"- 紀錄天數：{stats['days']} 天{extra}\n\n"
        "---\n\n"
        "## ✅ 建議\n"
        "- 維持 **高蛋白 (每公斤 1.6–2.0 g)** 與 **每週 2–3 次阻力訓練**  \n"
        "- 飲水 **≥ 3 L/天**（依活動量調整）  \n"
        "- 若每週下降 > 2.5 kg，建議微調熱量或與醫師討論  \n"
    )

    # KPI 目標與進度（每週）
    kpi = compute_weekly_kpi(stats)
    # 現況與達成度
    # 體重
    weight_delta = None
    if stats.get('start_weight_am') is not None and stats.get('end_weight_am') is not None:
        weight_delta = abs(stats['end_weight_am'] - stats['start_weight_am'])
    weight_bar = _progress_bar(
        current=stats.get('end_weight_am'),
        target_delta=abs(kpi.get('weight_target_end') - kpi.get('weight_start')) if kpi.get('weight_target_end') is not None and kpi.get('weight_start') is not None else None,
        achieved_delta=weight_delta if weight_delta is not None else 0,
        inverse=True
    )
    # 體脂率
    fat_delta = None
    if stats.get('start_fat_am') is not None and stats.get('end_fat_am') is not None:
        fat_delta = abs(stats['end_fat_am'] - stats['start_fat_am'])
    fat_bar = _progress_bar(
        current=stats.get('end_fat_am'),
        target_delta=abs(kpi.get('fat_pct_target_end') - kpi.get('fat_pct_start')) if kpi.get('fat_pct_target_end') is not None and kpi.get('fat_pct_start') is not None else None,
        achieved_delta=fat_delta if fat_delta is not None else 0,
        inverse=True
    )
    # 內臟脂肪
    vis_delta = None
    if stats.get('start_visceral_am') is not None and stats.get('end_visceral_am') is not None:
        vis_delta = abs(stats['end_visceral_am'] - stats['start_visceral_am'])
    vis_bar = _progress_bar(
        current=stats.get('end_visceral_am'),
        target_delta=abs(kpi.get('visceral_target_end') - kpi.get('visceral_start')) if kpi.get('visceral_target_end') is not None and kpi.get('visceral_start') is not None else None,
        achieved_delta=vis_delta if vis_delta is not None else 0,
        inverse=True
    )
    # 骨骼肌重量（保持/增加）
    musw_delta = None
    if stats.get('start_muscle_weight_am') is not None and stats.get('end_muscle_weight_am') is not None:
        musw_delta = stats['end_muscle_weight_am'] - stats['start_muscle_weight_am']
    musw_target = 0.0
    musw_bar = _progress_bar(
        current=stats.get('end_muscle_weight_am'),
        target_delta=musw_target if musw_target > 0 else 0.001,  # avoid zero division
        achieved_delta=max(0.0, musw_delta) if musw_delta is not None else 0.0,
        inverse=False
    )

    # 組成品質（最近28天：脂肪下降/體重下降）
    ratio, qd = compute_quality_ratio(wdf, days=28)
    if ratio is not None:
        label = "良好" if ratio >= 0.6 else ("普通" if ratio >= 0.4 else "需留意")
        md += f"\n---\n\n## 🧪 組成品質（近28天）\n\n- 脂肪/體重 下降比例：{ratio*100:.0f}%（{label}）  \n- 體重變化：-{qd['weight_drop']:.1f} kg，脂肪重量變化：-{qd['fat_drop']:.1f} kg（AM）  \n"

    md += f"\n---\n\n## 🎯 KPI 目標與進度 ({kpi_period_label})\n\n"
    md += "- 體重：目標 -0.8 kg  \n"
    if kpi.get('weight_start') is not None and kpi.get('weight_target_end') is not None:
        md += f"  - 由 {kpi['weight_start']:.1f} → 目標 {kpi['weight_target_end']:.1f} kg  | 進度 {weight_bar}  \n"
    md += "- 體脂率（AM）：目標 -0.4 個百分點  \n"
    if kpi.get('fat_pct_start') is not None and kpi.get('fat_pct_target_end') is not None:
        md += f"  - 由 {kpi['fat_pct_start']:.1f}% → 目標 {kpi['fat_pct_target_end']:.1f}%  | 進度 {fat_bar}  \n"
    md += "- 內臟脂肪（AM）：目標 -0.5  \n"
    if kpi.get('visceral_start') is not None and kpi.get('visceral_target_end') is not None:
        md += f"  - 由 {kpi['visceral_start']:.1f} → 目標 {kpi['visceral_target_end']:.1f}  | 進度 {vis_bar}  \n"
    if stats.get('start_muscle_weight_am') is not None and stats.get('end_muscle_weight_am') is not None:
        md += f"- 骨骼肌重量（AM）：目標 ≥ 持平  | 變化 {stats['end_muscle_weight_am']-stats['start_muscle_weight_am']:+.1f} kg  | 進度 {musw_bar}  \n"

    # 目標 ETA（近28天趨勢估算）
    if goals:
        gw = goals.get('weight_final') if isinstance(goals, dict) else None
        gf = goals.get('fat_pct_final') if isinstance(goals, dict) else None
        # 預設採用統一視窗 + 脂肪重量
        # 主要 ETA：脂肪重量對應目標（由目標體重與體脂率換算）
        fat_eta_line = ""
        try:
            # 換算目標脂肪重量（以體重與體脂率）
            if gw is not None and gf is not None and '早上脂肪重量 (kg)' in wdf.columns:
                target_fatkg = gw * gf / 100.0
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_fk = _compute_eta(wdf_all=wdf, wdf_slice=wdf, metric='fatkg', target=target_fatkg, scope=scope, method=method)
                if eta_fk:
                    fat_eta_line = f"- 脂肪重量達標 ETA：~{eta_fk['weeks']:.1f} 週（{eta_fk['date']}）  \n"
        except Exception:
            pass
        # 次要：體重、體脂率（若有意義才顯示）
        try:
            if gw is not None:
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_w = _compute_eta(wdf_all=wdf, wdf_slice=wdf, metric='weight', target=gw, scope=scope, method=method)
                if eta_w:
                    md += f"- 體重達標 ETA：~{eta_w['weeks']:.1f} 週（{eta_w['date']}）  \n"
            if gf is not None:
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_f = _compute_eta(wdf_all=wdf, wdf_slice=wdf, metric='fatpct', target=gf, scope=scope, method=method)
                if eta_f:
                    md += f"- 體脂率達標 ETA（AM）：~{eta_f['weeks']:.1f} 週（{eta_f['date']}）  \n"
        except Exception:
            pass
        if fat_eta_line:
            md += fat_eta_line

    # 本期分析與總結（自動）
    md += "\n---\n\n## 🧠 本期數據分析與總結\n\n"
    # 亮點
    if stats.get('delta_weight_am') is not None and stats['delta_weight_am'] < 0:
        md += f"- ✅ 體重：{abs(stats['delta_weight_am']):.1f} kg 下降（AM）\n"
    if stats.get('delta_fat_am') is not None and stats['delta_fat_am'] < 0:
        md += f"- ✅ 體脂率：{abs(stats['delta_fat_am']):.1f} 個百分點下降（AM）\n"
    if stats.get('delta_visceral_am') is not None and stats['delta_visceral_am'] < 0:
        md += f"- ✅ 內臟脂肪：{abs(stats['delta_visceral_am']):.1f} 降低（AM）\n"
    if stats.get('delta_muscle_am') is not None and stats['delta_muscle_am'] > 0:
        md += f"- ✅ 骨骼肌率：+{abs(stats['delta_muscle_am']):.1f} 個百分點（AM）\n"
    if stats.get('delta_fat_weight_am') is not None and stats['delta_fat_weight_am'] < 0:
        md += f"- ✅ 脂肪重量：-{abs(stats['delta_fat_weight_am']):.1f} kg（AM）\n"
    
    # 風險提示
    if stats.get('delta_muscle_weight_am') is not None and stats['delta_muscle_weight_am'] < 0:
        md += f"- ⚠️ 骨骼肌重量下降：{abs(stats['delta_muscle_weight_am']):.1f} kg，建議調整赤字與訓練恢復。\n"
    ratio, qd = compute_quality_ratio(wdf, days=28)
    if ratio is not None and ratio < 0.4:
        md += "- ⚠️ 組成品質偏低（脂肪/體重 < 40%），建議提高蛋白與阻力訓練，減少過大赤字。\n"

    # 下一步（簡短）
    md += "\n- 下一步：蛋白 1.8–2.2 g/kg、每週 3–4 次阻力訓練、穩定睡眠與步數，維持每週 -0.5～-0.8 kg。\n"

    # 寫入檔案
    with open(out_md_path, "w", encoding="utf-8") as f:
        f.write(md)

def make_summary_report(df, out_dir, prefix="summary", goals: dict | None = None, eta_config: dict | None = None, show_targets: bool = True):
    """產生從第一天到最新數據的總結報告"""
    df_sorted = df.sort_values("日期")
    
    # 計算整體統計
    stats = compute_stats(df_sorted)
    
    # 產生圖表（若有長期目標，亦可疊加目標線：用 is_week=True 的線性輔助）
    summary_kpi = None
    if goals and (goals.get('weight_final') or goals.get('fat_pct_final')):
        summary_kpi = {}
        # 以第一天為起點，目標為最終值；用全期間長度做線性參考線
        if goals.get('weight_final') is not None:
            summary_kpi['weight_start'] = df_sorted['早上體重 (kg)'].dropna().iloc[0] if not df_sorted['早上體重 (kg)'].dropna().empty else None
            summary_kpi['weight_target_end'] = goals['weight_final'] if summary_kpi['weight_start'] is not None else None
        if goals.get('fat_pct_final') is not None:
            start_fat = df_sorted['早上體脂 (%)'].dropna().iloc[0] if not df_sorted['早上體脂 (%)'].dropna().empty else None
            summary_kpi['fat_pct_start'] = start_fat
            summary_kpi['fat_pct_target_end'] = goals['fat_pct_final'] if start_fat is not None else None
    weight_png, bodyfat_png, visceral_png, muscle_png = make_charts(df_sorted, out_dir, prefix=prefix, kpi=summary_kpi, is_week=bool(summary_kpi), show_ma=True, show_targets=show_targets)
    
    # 計算週次
    total_days = len(df_sorted)
    total_weeks = (total_days + 6) // 7  # 向上取整
    
    # 產生表格 - 只顯示最近7天和第一天作對比
    recent_data = df_sorted.tail(7)
    first_day = df_sorted.iloc[0:1].copy()
    
    # 表格欄位
    table_cols = ["日期","早上體重 (kg)","晚上體重 (kg)","早上體脂 (%)","晚上體脂 (%)"]
    has_visceral = '早上內臟脂肪' in df_sorted.columns and '晚上內臟脂肪' in df_sorted.columns
    has_muscle = '早上骨骼肌 (%)' in df_sorted.columns and '晚上骨骼肌 (%)' in df_sorted.columns
    if has_visceral:
        table_cols.extend(["早上內臟脂肪","晚上內臟脂肪"])
    if has_muscle:
        table_cols.extend(["早上骨骼肌 (%)","晚上骨骼肌 (%)"])
    
    if len(df_sorted) <= 7:
        display_data = df_sorted[table_cols].copy()
    else:
        # 創建分隔行
        separator_dict = {"日期": ["..."], "早上體重 (kg)": ["..."], "晚上體重 (kg)": ["..."], 
                         "早上體脂 (%)": ["..."], "晚上體脂 (%)": ["..."]}
        if has_visceral:
            separator_dict["早上內臟脂肪"] = ["..."]
            separator_dict["晚上內臟脂肪"] = ["..."]
        if has_muscle:
            separator_dict["早上骨骼肌 (%)"] = ["..."]
            separator_dict["晚上骨骼肌 (%)"] = ["..."]
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
    if muscle_png:
        charts_section += f"![骨骼肌趨勢]({os.path.basename(muscle_png)})\n"
    # 組成品質（最近28天：脂肪下降/體重下降）
    ratio, qd = compute_quality_ratio(df_sorted, days=28)
    if ratio is not None:
        label = "良好" if ratio >= 0.6 else ("普通" if ratio >= 0.4 else "需留意")
        charts_section += (
            "\n## 🧪 組成品質（近28天）\n\n"
            f"- 脂肪/體重 下降比例：{ratio*100:.0f}%（{label}）  \n"
            f"- 體重變化：-{qd['weight_drop']:.1f} kg，脂肪重量變化：-{qd['fat_drop']:.1f} kg（AM）  \n\n"
            "---\n\n"
        )
    else:
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
    
    # 骨骼肌統計
    muscle_stats = ""
    if stats.get("avg_muscle_am") is not None:
        muscle_stats = (
            f"\n- **骨骼肌（AM）**：{_fmt(stats['start_muscle_am'], 1)}% → {_fmt(stats['end_muscle_am'], 1)}%  "
            f"(**{_fmt(stats['delta_muscle_am'], 1)}%**), 總平均 {stats['avg_muscle_am']:.1f}%  \n"
            f"- **骨骼肌（PM）**：{_fmt(stats['start_muscle_pm'], 1)}% → {_fmt(stats['end_muscle_pm'], 1)}%  "
            f"(**{_fmt(stats['delta_muscle_pm'], 1)}%**), 總平均 {stats['avg_muscle_pm']:.1f}%  \n"
            f"- **骨骼肌（AM+PM 平均）**：{stats['avg_muscle_all']:.1f}%  \n"
        )
    
    # 脂肪重量統計
    fat_weight_stats = ""
    if stats.get("avg_fat_weight_am") is not None:
        fat_weight_stats = (
            f"\n- **脂肪重量（AM）**：{_fmt(stats['start_fat_weight_am'], 1)} → {_fmt(stats['end_fat_weight_am'], 1)} kg  "
            f"(**{_fmt(stats['delta_fat_weight_am'], 1)} kg**), 總平均 {stats['avg_fat_weight_am']:.1f} kg  \n"
            f"- **脂肪重量（PM）**：{_fmt(stats['start_fat_weight_pm'], 1)} → {_fmt(stats['end_fat_weight_pm'], 1)} kg  "
            f"(**{_fmt(stats['delta_fat_weight_pm'], 1)} kg**), 總平均 {stats['avg_fat_weight_pm']:.1f} kg  \n"
            f"- **脂肪重量（AM+PM 平均）**：{stats['avg_fat_weight_all']:.1f} kg  \n"
        )
    
    # 骨骼肌重量統計
    muscle_weight_stats = ""
    if stats.get("avg_muscle_weight_am") is not None:
        muscle_weight_stats = (
            f"\n- **骨骼肌重量（AM）**：{_fmt(stats['start_muscle_weight_am'], 1)} → {_fmt(stats['end_muscle_weight_am'], 1)} kg  "
            f"(**{_fmt(stats['delta_muscle_weight_am'], 1)} kg**), 總平均 {stats['avg_muscle_weight_am']:.1f} kg  \n"
            f"- **骨骼肌重量（PM）**：{_fmt(stats['start_muscle_weight_pm'], 1)} → {_fmt(stats['end_muscle_weight_pm'], 1)} kg  "
            f"(**{_fmt(stats['delta_muscle_weight_pm'], 1)} kg**), 總平均 {stats['avg_muscle_weight_pm']:.1f} kg  \n"
            f"- **骨骼肌重量（AM+PM 平均）**：{stats['avg_muscle_weight_all']:.1f} kg  \n"
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
        f"{visceral_stats}"
        f"{muscle_stats}"
        f"{fat_weight_stats}"
        f"{muscle_weight_stats}\n"
        f"- **追蹤天數**：{stats['days']} 天{extra}{weekly_analysis}\n\n"
        "---\n\n"
        "## 🎯 重點成果\n\n"
    )
    
    # 若有長期目標，加入目標達成進度（以 AM 值為主）
    if goals and (goals.get('weight_final') is not None or goals.get('fat_pct_final') is not None):
        md += "### 🎯 長期目標進度\n"
        if goals.get('weight_final') is not None and stats.get('end_weight_am') is not None:
            start_w = stats.get('start_weight_am')
            end_w = stats.get('end_weight_am')
            goal_w = goals['weight_final']
            total_drop = (start_w - goal_w) if (start_w is not None and goal_w is not None) else None
            achieved = (start_w - end_w) if (start_w is not None and end_w is not None) else None
            w_bar = _progress_bar(current=end_w, target_delta=abs(total_drop) if total_drop is not None else None, achieved_delta=abs(achieved) if achieved is not None else 0, inverse=True)
            md += f"- 體重目標：{start_w:.1f} → {goal_w:.1f} kg  | 目前 {end_w:.1f} kg  | 進度 {w_bar}  \n"
        if goals.get('fat_pct_final') is not None and stats.get('end_fat_am') is not None:
            start_f = stats.get('start_fat_am')
            end_f = stats.get('end_fat_am')
            goal_f = goals['fat_pct_final']
            total_drop = (start_f - goal_f) if (start_f is not None and goal_f is not None) else None
            achieved = (start_f - end_f) if (start_f is not None and end_f is not None) else None
            f_bar = _progress_bar(current=end_f, target_delta=abs(total_drop) if total_drop is not None else None, achieved_delta=abs(achieved) if achieved is not None else 0, inverse=True)
            md += f"- 體脂率目標（AM）：{start_f:.1f}% → {goal_f:.1f}%  | 目前 {end_f:.1f}%  | 進度 {f_bar}  \n"
        # 目標 ETA（近28天趨勢估算）
        # 動態方法標籤
        _method = (eta_config or {}).get('method', 'regress28')
        _method_label = {
            'endpoint_all': '首末端點（全期間）',
            'regress_all': '線性回歸（全期間）',
            'regress28': '線性回歸（近28天）',
            'endpoint28': '首末端點（近28天）',
        }.get(_method, '趨勢估算')
        md += f"\n#### ⏱️ 目標 ETA（{_method_label}）\n"
        try:
            gw = goals.get('weight_final'); gf = goals.get('fat_pct_final')
            # 初始化旗標以便必要時提供友善提示
            printed_any = False
            if gw is not None and gf is not None:
                target_fatkg = gw * gf / 100.0
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_fk = _compute_eta(wdf_all=df_sorted, wdf_slice=df_sorted, metric='fatkg', target=target_fatkg, scope=scope, method=method)
                if eta_fk:
                    md += f"- 脂肪重量達標 ETA：~{eta_fk['weeks']:.1f} 週（{eta_fk['date']}）  \n"
                    printed_any = True
                else:
                    md += f"- 脂肪重量達標 ETA：暫無穩定趨勢，無法估算（{_method_label}）  \n"
                    printed_any = True
            if gw is not None:
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_w = _compute_eta(wdf_all=df_sorted, wdf_slice=df_sorted, metric='weight', target=gw, scope=scope, method=method)
                if eta_w:
                    md += f"- 體重達標 ETA：~{eta_w['weeks']:.1f} 週（{eta_w['date']}）  \n"
                    printed_any = True
            if gf is not None:
                scope = (eta_config or {}).get('scope', 'global')
                method = (eta_config or {}).get('method', 'regress28')
                eta_f = _compute_eta(wdf_all=df_sorted, wdf_slice=df_sorted, metric='fatpct', target=gf, scope=scope, method=method)
                if eta_f:
                    md += f"- 體脂率達標 ETA（AM）：~{eta_f['weeks']:.1f} 週（{eta_f['date']}）  \n"
                    printed_any = True
            if not printed_any:
                md += f"- 資料趨勢不足（{_method_label}），暫無 ETA 可供參考  \n"
        except Exception:
            md += "- ETA 計算發生例外，暫無 ETA 可供參考  \n"
    
    # 成果分析
    if stats['delta_weight_am'] and stats['delta_weight_am'] < 0:
        md += f"✅ **體重減少**：在 {total_days} 天內減重 {abs(stats['delta_weight_am']):.1f} kg（早上測量）  \n"
    if stats['delta_fat_pm'] and stats['delta_fat_pm'] < 0:
        md += f"✅ **體脂下降**：體脂率降低 {abs(stats['delta_fat_pm']):.1f}%（晚上測量）  \n"
    if stats.get('delta_visceral_am') and stats['delta_visceral_am'] < 0:
        md += f"✅ **內臟脂肪改善**：內臟脂肪程度降低 {abs(stats['delta_visceral_am']):.1f}（早上測量）  \n"
    if stats.get('delta_muscle_am') and stats['delta_muscle_am'] > 0:
        md += f"✅ **骨骼肌增加**：骨骼肌率提升 {abs(stats['delta_muscle_am']):.1f}%（早上測量）  \n"
    if stats.get('delta_fat_weight_am') and stats['delta_fat_weight_am'] < 0:
        md += f"✅ **脂肪重量減少**：減少 {abs(stats['delta_fat_weight_am']):.1f} kg 脂肪（早上測量）  \n"
    if stats.get('delta_muscle_weight_am') and stats['delta_muscle_weight_am'] > 0:
        md += f"✅ **骨骼肌重量增加**：增加 {abs(stats['delta_muscle_weight_am']):.1f} kg 骨骼肌（早上測量）  \n"
    
    md += "\n## ✅ 持續建議\n"
    md += "- 維持 **高蛋白 (每公斤 1.6–2.0 g)** 與 **每週 2–3 次阻力訓練**  \n"
    md += "- 飲水 **≥ 3 L/天**（依活動量調整）  \n"
    md += "- 持續監測體重與體脂變化，建議保持每週穩定減重  \n"
    md += "- 如有任何異常變化，建議諮詢專業醫師  \n"
    
    return md, weight_png, bodyfat_png, visceral_png, muscle_png

def main():
    p = argparse.ArgumentParser(description="以週五為起始的自訂週期，從 master 產生 Excel + Markdown + 圖表（支援 CSV/Excel 格式）")
    p.add_argument("master", nargs="?", default="BodyComposition_202507-202510.csv", help="主檔（CSV 或 Excel 格式）")
    p.add_argument("--sheet", default=None, help="工作表名稱（僅用於 Excel，預設先嘗試 'Daily Log'，再退回第一個工作表）")
    p.add_argument("--header-row", type=int, default=0, help="欄位標題所在的列索引（僅用於 Excel，0=第一列）")
    p.add_argument("--anchor-date", default="2025-08-15", help="每週起始的對齊基準日（週四），例如 2025-08-15")
    p.add_argument("--week-index", type=int, default=None, help="第幾週（以 anchor-date 為第1週起算）；未提供則取最後一週")
    p.add_argument("--out-root", default=".", help="輸出根目錄（會在裡面建立 weekly/ 與 reports/）")
    p.add_argument("--summary", action="store_true", help="產生從第一天到最新數據的總結報告")
    p.add_argument("--goal-weight", type=float, default=79, help="最終目標體重 (kg)，用於總結報告的目標與進度（預設：79）")
    p.add_argument("--goal-fat-pct", type=float, default=12, help="最終目標體脂率 (%)，用於總結報告的目標與進度（預設：12）")
    p.add_argument("--monthly", nargs="?", const="latest", help="產生某月份的月度報告（YYYY-MM，不帶值則取最新月份）")
    p.add_argument("--eta-scope", choices=["global","local"], default="global", help="ETA 計算視窗：global=用全資料最後日回推28天；local=用當前報告子集最後日回推28天")
    p.add_argument("--eta-metric", choices=["fatkg","weight","fatpct"], default="fatkg", help="ETA 主要估算指標：脂肪重量、體重或體脂率")
    p.add_argument("--eta-method", choices=["regress28","endpoint_all","regress_all","endpoint28"], default="endpoint_all", help="ETA 估算方法：regress28=近28天回歸、endpoint_all=首末端點、regress_all=全期間回歸、endpoint28=近28天端點（預設：endpoint_all）")
    # 圖表目標線：預設不顯示，使用 --show-target-lines 可打開
    group = p.add_mutually_exclusive_group()
    group.add_argument("--no-target-lines", action="store_true", help="不在圖表上繪製目標參考線（預設）")
    group.add_argument("--show-target-lines", action="store_true", help="在圖表上繪製目標參考線")
    args = p.parse_args()

    # 預設：不畫目標線（若未提供兩個旗標，維持預設不顯示）
    if not args.no_target_lines and not args.show_target_lines:
        args.no_target_lines = True

    df = read_daily_log(args.master, sheet_name=args.sheet, header_row=args.header_row)

    if args.summary:
        # 產生總結報告
        reports_dir = os.path.join(args.out_root, "reports")
        summary_dir = os.path.join(reports_dir, "summary")
        ensure_dirs(summary_dir)
        
        chart_show_targets = True if args.show_target_lines else (not args.no_target_lines)
        summary_md, weight_png, bodyfat_png, visceral_png, muscle_png = make_summary_report(
            df, summary_dir, goals={
                'weight_final': args.goal_weight,
                'fat_pct_final': args.goal_fat_pct,
            }, eta_config={'scope': args.eta_scope, 'method': args.eta_method}, show_targets=chart_show_targets
        )
        summary_md_path = os.path.join(summary_dir, "overall_summary_report.md")
        
        with open(summary_md_path, "w", encoding="utf-8") as f:
            f.write(summary_md)
        
        print("✅ 總結報告已完成輸出")
        print("Summary MD :", summary_md_path)
        charts_list = [weight_png, bodyfat_png]
        if visceral_png:
            charts_list.append(visceral_png)
        if muscle_png:
            charts_list.append(muscle_png)
        print("Charts     :", " ".join(charts_list))
        return

    # 月報模式
    if args.monthly is not None:
        reports_dir = os.path.join(args.out_root, "reports")
        ym = None if args.monthly == "latest" else args.monthly
        wdf, ym_tag, start_date, end_date = pick_month(df, ym)
        month_dir = os.path.join(reports_dir, "monthly", ym_tag)
        ensure_dirs(month_dir)

        # 以每週目標為基礎，放大至本月天數/週數
        stats = compute_stats(wdf)
        weeks = max(1, (len(wdf) + 6) // 7)
        base_kpi = compute_weekly_kpi(stats)
        # 放大：體重 0.8*weeks、體脂 0.4*weeks、內臟 0.5*weeks
        month_kpi = {}
        if base_kpi.get('weight_start') is not None:
            month_kpi['weight_start'] = base_kpi['weight_start']
            month_kpi['weight_target_end'] = base_kpi['weight_start'] - 0.8 * weeks
        if base_kpi.get('fat_pct_start') is not None:
            month_kpi['fat_pct_start'] = base_kpi['fat_pct_start']
            month_kpi['fat_pct_target_end'] = max(base_kpi['fat_pct_start'] - 0.4 * weeks, 0)
        if base_kpi.get('visceral_start') is not None:
            month_kpi['visceral_start'] = base_kpi['visceral_start']
            month_kpi['visceral_target_end'] = max(base_kpi['visceral_start'] - 0.5 * weeks, 0)
        if base_kpi.get('muscle_pct_floor') is not None:
            month_kpi['muscle_pct_floor'] = base_kpi['muscle_pct_floor']
        if base_kpi.get('muscle_weight_start') is not None:
            month_kpi['muscle_weight_start'] = base_kpi['muscle_weight_start']
            month_kpi['muscle_weight_target_end'] = base_kpi['muscle_weight_start']

        # 圖表（加上月度目標線）
        chart_show_targets = True if args.show_target_lines else (not args.no_target_lines)
        weight_png, bodyfat_png, visceral_png, muscle_png = make_charts(wdf, month_dir, prefix=f"{ym_tag}", kpi=month_kpi, is_week=True, show_ma=True, show_targets=chart_show_targets)

        # 產出 MD（沿用週報版樣式，標題與文案換成月報）
        md_path = os.path.join(month_dir, f"{ym_tag}_monthly_report.md")
        # 借用 make_markdown：顯示同樣的統計文字與 KPI 區塊
        # 月報：帶入長期目標，顯示 ETA（若 CLI 有提供）
        month_goals = {
            'weight_final': args.goal_weight,
            'fat_pct_final': args.goal_fat_pct,
        }
        if month_goals['weight_final'] is None and month_goals['fat_pct_final'] is None:
            month_goals = None
        make_markdown(wdf, stats, weight_png, bodyfat_png, visceral_png, muscle_png, md_path, f"{ym_tag} 月報", start_date, end_date, kpi_period_label="本月", goals=month_goals, eta_config={'scope': args.eta_scope, 'method': args.eta_method})
        print("✅ 月度報告已完成輸出")
        print("Monthly MD:", md_path)
        return

    wdf, week_tag, start_date, end_date = pick_custom_week(df, args.anchor_date, args.week_index)

    weekly_dir = os.path.join(args.out_root, "weekly")
    reports_dir = os.path.join(args.out_root, "reports")
    week_reports_dir = os.path.join(reports_dir, week_tag)  # 在 reports 下建立週期子資料夾
    ensure_dirs(weekly_dir); ensure_dirs(week_reports_dir)

    weekly_xlsx = os.path.join(weekly_dir, f"{week_tag}_weight_tracking.xlsx")
    save_weekly_excel(wdf, weekly_xlsx)

    # 每週 KPI
    stats = compute_stats(wdf)
    kpi = compute_weekly_kpi(stats)

    chart_show_targets = True if args.show_target_lines else (not args.no_target_lines)
    weight_png, bodyfat_png, visceral_png, muscle_png = make_charts(wdf, week_reports_dir, prefix=week_tag, kpi=kpi, is_week=True, show_ma=True, show_targets=chart_show_targets)

    weekly_md = os.path.join(week_reports_dir, f"{week_tag}_weekly_report.md")
    # 將長期目標（若 CLI 有提供）帶入週報，顯示 ETA
    weekly_goals = {
        'weight_final': args.goal_weight,
        'fat_pct_final': args.goal_fat_pct,
    }
    if weekly_goals['weight_final'] is None and weekly_goals['fat_pct_final'] is None:
        weekly_goals = None
    make_markdown(wdf, stats, weight_png, bodyfat_png, visceral_png, muscle_png, weekly_md, week_tag, start_date, end_date, kpi_period_label="本週", goals=weekly_goals, eta_config={'scope': args.eta_scope, 'method': args.eta_method})

    print("✅ 已完成輸出")
    print("Weekly Excel:", weekly_xlsx)
    print("Report MD   :", weekly_md)
    charts_list = [weight_png, bodyfat_png]
    if visceral_png:
        charts_list.append(visceral_png)
    if muscle_png:
        charts_list.append(muscle_png)
    print("Charts      :", " ".join(charts_list))

if __name__ == "__main__":
    main()


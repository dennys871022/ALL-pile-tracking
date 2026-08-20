import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from shapely.geometry import Polygon
import os
import tempfile
from datetime import datetime, timedelta, date
from streamlit_gsheets import GSheetsConnection
import re

st.set_page_config(page_title="後台管理端", layout="wide")
st.title("🚧 CDC土方管理系統 ")

st.sidebar.markdown("### 🔒 系統權限")
pwd = st.sidebar.text_input("輸入管理員密碼解鎖編輯模式", type="password")
if pwd == "34561297":
    demo_mode = False
    st.sidebar.success("🔓 已解鎖管理員模式")
else:
    demo_mode = True
    if pwd:
        st.sidebar.error("❌ 密碼錯誤")
    else:
        st.sidebar.caption("👀 目前為訪客模式 (唯讀沙盒)")

SHEET_URL = "https://docs.google.com/spreadsheets/d/1y3Qnlx9qFwV6S6pyFTsT4rlXP_Tb8qd9tNhRBTjBHao/edit"

if 'canvas_key_counter' not in st.session_state:
    st.session_state['canvas_key_counter'] = 0

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"資料庫連線失敗：{e}")
    st.stop()

@st.cache_data(ttl=300)
def fetch_gsheet_data(sheet_name):
    return conn.read(spreadsheet=SHEET_URL, worksheet=sheet_name, ttl=0)

def load_sheet_data(sheet_name):
    try:
        df = fetch_gsheet_data(sheet_name)
        df = df.dropna(how='all')
        if not df.empty:
            st.session_state[f"cache_{sheet_name}"] = df.copy()
        return df
    except Exception:
        if f"cache_{sheet_name}" in st.session_state:
            return st.session_state[f"cache_{sheet_name}"]
        return pd.DataFrame()

def save_sheet_data(sheet_name, df):
    if demo_mode:
        st.session_state[f"cache_{sheet_name}"] = df.copy()
        st.toast(f"👀 訪客模式：已模擬儲存【{sheet_name}】至網頁暫存，雲端資料庫未變動。")
        return True

    try:
        conn.update(spreadsheet=SHEET_URL, worksheet=sheet_name, data=df)
        st.cache_data.clear()
        st.session_state[f"cache_{sheet_name}"] = df.copy()
        return True
    except Exception as e:
        st.error(f"寫入分頁 `{sheet_name}` 失敗：{e}")
        return False

if st.sidebar.button("🔄 強制同步雲端最新資料", use_container_width=True):
    for sheet in ["grid_zones", "dispatch_logs", "manifest_settings", "manifest_delivery", "stage_settings", "stage_daily_notes"]:
        if f"cache_{sheet}" in st.session_state:
            del st.session_state[f"cache_{sheet}"]
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("### 🎯 目前作業階段")
STAGE_OPTIONS = ["開挖前土方", "第1階段 (第1挖)", "第2階段 (第2挖)", "第3階段 (第3挖)", "第4階段 (第4挖)"]
if "global_stage_choice" not in st.session_state:
    st.session_state["global_stage_choice"] = STAGE_OPTIONS[0]
global_stage_choice = st.sidebar.selectbox(
    "影響：階段管控頁 / 出土儀表板PDF / 單階段地圖",
    STAGE_OPTIONS,
    key="global_stage_choice",
)

st.sidebar.markdown("### 各區開挖 GL 高程設定")
base_x_input = st.sidebar.number_input("1軸與A軸交點 X", value=-274766.4, format="%.2f")
base_y_input = st.sidebar.number_input("1軸與A軸交點 Y", value=-24009.49, format="%.2f")
scale_option = st.sidebar.selectbox("CAD圖資單位", ["公分 (除以100)", "公尺 (不轉換)", "公釐 (除以1000)"])
scale_factor = 100 if "公分" in scale_option else (1000 if "公釐" in scale_option else 1)
current_gl = st.sidebar.number_input("現地 GL 高程增減 (m)", value=0.0, step=0.1)

gl_admin_input = st.sidebar.text_input("行政棟區域 GL高程 (4挖)", "2.5, 4.45, 7.85, 9.9")
gl_lab_input = st.sidebar.text_input("實驗棟區域 GL高程 (4挖)", "2.5, 4.45, 7.85, 11.4")
gl_bc_input = st.sidebar.text_input("滯洪池BC區 GL高程 (2挖)", "1.5, 7.6")
gl_a_input = st.sidebar.text_input("滯洪池A區 GL高程 (2挖)", "2.0, 7.85")

def get_thickness_from_gl(gl_str, gl_offset):
    try:
        gl_list = [float(x.strip()) for x in gl_str.split(",")]
        thickness = []
        for i in range(len(gl_list)):
            if i == 0:
                thickness.append(max(0.0, gl_list[i] + gl_offset))
            else:
                thickness.append(max(0.0, gl_list[i] - gl_list[i-1]))
        return thickness
    except Exception:
        return []

# ============================================================================
# 新增：階段總覽計算（原本寫死在 tab_stage 裡，現在抽出來讓 PDF 報表也能共用）
# ============================================================================
def _migrate_stage_settings_schema(df):
    """
    相容舊版雲端資料：stage_settings 分頁舊欄位是「預計施作工期」（天數），
    新版改成「預計結束日期」。如果讀到的資料還是舊欄位，這裡自動換算補上新欄位，
    避免舊資料在新版程式裡直接 KeyError。
    """
    df = df.copy()
    if "預計結束日期" not in df.columns:
        if "預計施作工期" in df.columns and "預計開始時間" in df.columns:
            def _calc_end(row):
                try:
                    start = pd.to_datetime(row["預計開始時間"])
                    days = pd.to_numeric(row["預計施作工期"], errors='coerce')
                    if pd.isna(days):
                        days = 20
                    return (start + pd.Timedelta(days=int(days))).strftime("%Y-%m-%d")
                except Exception:
                    return None
            df["預計結束日期"] = df.apply(_calc_end, axis=1)
        else:
            df["預計結束日期"] = None
    return df


def compute_stage_overview(stage_choice, df_results, override_settings_row=None, as_of_date=None):
    """
    計算單一階段（開挖前土方 / 第1~4挖）的總覽數據，回傳 dict。
    給 tab_stage（畫面顯示）與 tab_stats 的 PDF 匯出共用，確保兩邊數字一致。
    override_settings_row: 若提供（例如畫面上尚未儲存的 data_editor 編輯值），
    優先使用它來計算，取代從雲端讀回的已儲存設定，維持「編輯即時預覽」的效果。
    as_of_date: 若提供，代表「查詢過往報表」——把所有「今天」相關的計算都改成以這個日期為準
    （該日的出土量、累積量、剩餘量、完成百分比等），呈現「截至那一天」的歷史快照。
    預設為 None，等同今天（維持原本即時總覽的行為）。
    """
    tw_today_ = (datetime.utcnow() + timedelta(hours=8)).date()
    query_date_ = as_of_date if as_of_date is not None else tw_today_

    df_stage_map_ = df_results.copy() if df_results is not None and not df_results.empty else pd.DataFrame()
    if "第1挖" in stage_choice:
        target_col_ = '第1挖方量(m³)'
    elif "第2挖" in stage_choice:
        target_col_ = '第2挖方量(m³)'
    elif "第3挖" in stage_choice:
        target_col_ = '第3挖方量(m³)'
        if not df_stage_map_.empty:
            df_stage_map_ = df_stage_map_[~df_stage_map_['分區代號'].str.contains("滯")]
    elif "第4挖" in stage_choice:
        target_col_ = '第4挖方量(m³)'
        if not df_stage_map_.empty:
            df_stage_map_ = df_stage_map_[~df_stage_map_['分區代號'].str.contains("滯")]
    else:
        target_col_ = None

    default_est_vol_ = df_stage_map_[target_col_].sum() if target_col_ and not df_stage_map_.empty else 0

    df_stage_set_ = load_sheet_data("stage_settings")
    if df_stage_set_.empty or "階段名稱" not in df_stage_set_.columns:
        df_stage_set_ = pd.DataFrame(columns=["階段名稱", "預計開始時間", "預計結束日期", "預估土方量(鬆方)", "單車預設實方"])

    current_set_ = df_stage_set_[df_stage_set_["階段名稱"] == stage_choice]
    if override_settings_row is not None:
        s_row_ = override_settings_row
    elif current_set_.empty:
        s_row_ = pd.Series({
            "階段名稱": stage_choice,
            "預計開始時間": tw_today_.strftime("%Y-%m-%d"),
            "預計結束日期": (tw_today_ + timedelta(days=20)).strftime("%Y-%m-%d"),
            "預估土方量(鬆方)": default_est_vol_,
            "單車預設實方": 12.0
        })
    else:
        s_row_ = current_set_.iloc[0]

    est_start_str_ = s_row_.get("預計開始時間", str(tw_today_))
    est_end_str_ = s_row_.get("預計結束日期", None)
    est_vol_ = pd.to_numeric(s_row_.get("預估土方量(鬆方)", 0), errors='coerce')
    vol_per_truck_ = pd.to_numeric(s_row_.get("單車預設實方", 12.0), errors='coerce')

    try:
        est_start_ = pd.to_datetime(est_start_str_).date()
    except Exception:
        est_start_ = tw_today_

    try:
        est_end_ = pd.to_datetime(est_end_str_).date() if est_end_str_ not in (None, "", "NaT") else est_start_ + timedelta(days=20)
    except Exception:
        est_end_ = est_start_ + timedelta(days=20)

    # 預計施作工期改由系統自動算：結束日期 - 開始日期（至少算1天，避免除以0）
    est_days_ = max((est_end_ - est_start_).days, 1)
    est_daily_vol_ = est_vol_ / est_days_ if est_days_ > 0 else 0

    df_logs_ = load_sheet_data("dispatch_logs")
    if not df_logs_.empty and "日期" in df_logs_.columns:
        valid_logs_ = df_logs_.copy()
        valid_logs_['載運方量(m³)'] = pd.to_numeric(valid_logs_.get('載運方量(m³)', vol_per_truck_), errors='coerce').fillna(vol_per_truck_)
        daily_stats_ = valid_logs_.groupby('日期').agg(
            實際車次=('車頭車號', 'count'),
            當日運棄量=('載運方量(m³)', 'sum')
        ).reset_index()
        daily_stats_ = daily_stats_.sort_values('日期')
    else:
        daily_stats_ = pd.DataFrame(columns=['日期', '實際車次', '當日運棄量'])

    df_daily_notes_ = load_sheet_data("stage_daily_notes")
    if df_daily_notes_.empty:
        df_daily_notes_ = pd.DataFrame(columns=['階段名稱', '日期', '內控預計車次', '實際車次', '差異', '當日運棄量', '累計運棄量', '剩餘土方量', '備註', '計入工期'])
    else:
        if '階段名稱' not in df_daily_notes_.columns:
            df_daily_notes_.insert(0, '階段名稱', stage_choice)
        if '剩餘土方量' not in df_daily_notes_.columns:
            df_daily_notes_['剩餘土方量'] = np.nan
        if '計入工期' not in df_daily_notes_.columns:
            df_daily_notes_['計入工期'] = np.nan

    curr_notes_ = df_daily_notes_[df_daily_notes_['階段名稱'] == stage_choice].copy() if not df_daily_notes_.empty else pd.DataFrame()

    max_log_date_ = pd.to_datetime(daily_stats_['日期'].max()).date() if not daily_stats_.empty else tw_today_
    end_date_for_range_ = max(tw_today_, max_log_date_)
    date_range_ = [d.strftime("%Y-%m-%d") for d in pd.date_range(est_start_, end_date_for_range_)]
    df_range_ = pd.DataFrame({"日期": date_range_})

    if not daily_stats_.empty:
        df_range_ = pd.merge(df_range_, daily_stats_, on="日期", how="left").fillna({"實際車次": 0, "當日運棄量": 0})
    else:
        df_range_['實際車次'] = 0
        df_range_['當日運棄量'] = 0

    df_range_['累計車次'] = df_range_['實際車次'].cumsum()
    df_range_['累計運棄量'] = df_range_['當日運棄量'].cumsum()

    if not curr_notes_.empty and '備註' in curr_notes_.columns:
        merge_cols_ = ['日期', '內控預計車次', '備註']
        if '計入工期' in curr_notes_.columns:
            merge_cols_.append('計入工期')
        merge_notes_ = curr_notes_[merge_cols_].drop_duplicates('日期')
        if '計入工期' in merge_notes_.columns:
            merge_notes_ = merge_notes_.rename(columns={'計入工期': '計入工期_saved'})
        df_range_ = pd.merge(df_range_, merge_notes_, on="日期", how="left")
    else:
        df_range_['內控預計車次'] = np.nan
        df_range_['備註'] = ""
        df_range_['計入工期_saved'] = np.nan

    default_daily_trips_ = round((est_vol_ / vol_per_truck_) / est_days_) if est_days_ > 0 and vol_per_truck_ > 0 else 0
    df_range_['內控預計車次'] = pd.to_numeric(df_range_['內控預計車次'], errors='coerce').fillna(default_daily_trips_)
    df_range_['差異'] = df_range_['實際車次'] - df_range_['內控預計車次']
    df_range_['備註'] = df_range_['備註'].fillna("")
    df_range_['剩餘土方量'] = (est_vol_ - df_range_['累計運棄量']).clip(lower=0)

    # 計入工期：有出土(>0)一律算，沒出土(=0)預設不算，除非該日被手動勾選為特例（存在 stage_daily_notes 裡）
    if '計入工期_saved' not in df_range_.columns:
        df_range_['計入工期_saved'] = np.nan
    df_range_['計入工期'] = df_range_['當日運棄量'] > 0
    _saved_mask = df_range_['計入工期_saved'].notna()
    df_range_.loc[_saved_mask, '計入工期'] = df_range_.loc[_saved_mask, '計入工期_saved'].astype(bool)
    df_range_.loc[df_range_['當日運棄量'] > 0, '計入工期'] = True  # 有出土一律算，不受任何手動設定影響
    df_range_ = df_range_.drop(columns=['計入工期_saved'])

    # 累積差異：只累加「有顯示的日期」（計入工期=True）的差異，被隱藏的0出土日不會貢獻差異值，
    # 這樣畫面上兩個相鄰可見日期之間的累積差異變化，才會等於中間那筆「差異」的值，數字對得起來
    df_range_['累積差異'] = df_range_['差異'].where(df_range_['計入工期'], 0).cumsum()

    actual_start_date_ = est_start_
    # 目前作業工期改用「有計入工期的天數」，不是單純日曆天數，
    # 這樣沒出土又沒被標記為特例的日期就不會拖累平均出土功率/剩餘天數的計算
    today_str_ = query_date_.strftime("%Y-%m-%d")
    _work_days_mask = (df_range_['日期'] <= today_str_) & (df_range_['計入工期'] == True)
    current_work_days_ = int(_work_days_mask.sum())
    today_row_ = df_range_[df_range_['日期'] == today_str_]
    today_vol_ = today_row_['當日運棄量'].sum() if not today_row_.empty else 0

    # cum_vol / cum_trips 用「截至查詢日那一列」已經算好的累計欄位，
    # 而不是整個 df_range_ 加總，這樣查詢過去日期時才不會把查詢日之後的資料也算進去
    if not today_row_.empty:
        cum_vol_ = today_row_['累計運棄量'].iloc[0]
        cum_trips_ = today_row_['累計車次'].iloc[0]
    else:
        past_rows_ = df_range_[df_range_['日期'] <= today_str_]
        cum_vol_ = past_rows_['當日運棄量'].sum() if not past_rows_.empty else 0
        cum_trips_ = past_rows_['實際車次'].sum() if not past_rows_.empty else 0

    avg_vol_per_day_ = cum_vol_ / current_work_days_ if current_work_days_ > 0 else 0
    remain_vol_ = max(0, est_vol_ - cum_vol_)
    remain_days_ = round(remain_vol_ / avg_vol_per_day_) if avg_vol_per_day_ > 0 else 0
    est_completion_date_ = query_date_ + timedelta(days=remain_days_)
    diff_days_ = (est_completion_date_ - est_end_).days
    percent_done_ = round((cum_vol_ / est_vol_) * 100) if est_vol_ > 0 else 0

    avg_trips_per_day_ = cum_trips_ / current_work_days_ if current_work_days_ > 0 else 0

    return {
        "stage_choice": stage_choice,
        "today": query_date_,
        "is_historical": as_of_date is not None,
        "current_work_days": current_work_days_,
        "est_start": est_start_,
        "actual_start_date": actual_start_date_,
        "est_days": est_days_,
        "remain_days": remain_days_,
        "est_end": est_end_,
        "est_completion_date": est_completion_date_,
        "diff_days": diff_days_,
        "est_vol": est_vol_,
        "today_vol": today_vol_,
        "est_daily_vol": est_daily_vol_,
        "cum_vol": cum_vol_,
        "avg_vol_per_day": avg_vol_per_day_,
        "cum_trips": cum_trips_,
        "avg_trips_per_day": avg_trips_per_day_,
        "remain_vol": remain_vol_,
        "percent_done": percent_done_,
        "df_range": df_range_,
        "df_daily_notes": df_daily_notes_,
        "df_stage_map": df_stage_map_,
        "target_col": target_col_,
        "vol_per_truck": vol_per_truck_,
        "est_vol_default": default_est_vol_,
    }


def build_stage_overview_html(overview):
    """把 compute_stage_overview() 回傳的 dict 畫成紫色階段總覽表格的HTML字串（給tab_stage即時總覽與歷史查詢共用）。"""
    html_table = f"""
    <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; text-align: right;">
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">今天日期</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['today']}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">目前作業工期</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['current_work_days']}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">預計開始時間</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_start']}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">實際開始時間</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['actual_start_date']}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">預計施作工期</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_days']}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">推估剩餘天數</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['remain_days']}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">預計完成日期</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_end']}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">推估完成日期</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_completion_date']}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">差異 (天)</td>
            <td style="border: 1px solid #ccc; padding: 8px; background-color: #f39c12; color: white;">{overview['diff_days']}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">預估土方量(鬆方)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_vol']:,.1f}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">本日出土量(m³)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['today_vol']:,.1f}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">預估每日出土量(m³)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['est_daily_vol']:,.1f}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">累積出土數量(m³)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['cum_vol']:,.1f}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">平均出土功率(m³/天)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['avg_vol_per_day']:,.1f}</td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">剩餘土方量(m³)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['remain_vol']:,.1f}</td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">平均出土功率(台/天)</td>
            <td style="border: 1px solid #ccc; padding: 8px;">{overview['avg_trips_per_day']:,.1f}</td>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
        </tr>
        <tr>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
            <td style="border: 1px solid #ccc; padding: 8px;"></td>
            <td style="border: 1px solid #ccc; padding: 8px; text-align: left; background-color: #f9f9f9;">完成百分比</td>
            <td style="border: 1px solid #ccc; padding: 8px;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <div style="flex: 1; background-color: #e8e0ee; border-radius: 4px; overflow: hidden; height: 14px;">
                        <div style="width: {min(max(overview['percent_done'], 0), 100)}%; background-color: #8e44ad; height: 100%;"></div>
                    </div>
                    <span style="color: #8e44ad; font-weight: bold; white-space: nowrap;">{overview['percent_done']}%</span>
                </div>
            </td>
        </tr>
    </table>
    """

    return html_table


def get_stage_index(stage_choice):
    """回傳 0~3 代表第1~4挖；開挖前土方回傳 None（因為它不是以分區門檻定義的階段）。"""
    if "第1挖" in stage_choice:
        return 0
    if "第2挖" in stage_choice:
        return 1
    if "第3挖" in stage_choice:
        return 2
    if "第4挖" in stage_choice:
        return 3
    return None


def _upsert_daily_notes_field(df_daily_notes, stage_choice, date_str, field, value):
    """
    只更新 stage_daily_notes 裡「某一天、某一個欄位」的值，其餘欄位（包含其他已存的手動設定）
    完全不動。給「0出土日期管理」跟其他只想改單一欄位的地方用，避免整列/整段覆寫。
    """
    df_daily_notes = df_daily_notes.copy()
    mask = (df_daily_notes['階段名稱'] == stage_choice) & (df_daily_notes['日期'].astype(str) == str(date_str))
    if mask.any():
        df_daily_notes.loc[mask, field] = value
    else:
        new_row = {c: np.nan for c in df_daily_notes.columns}
        new_row['階段名稱'] = stage_choice
        new_row['日期'] = date_str
        new_row[field] = value
        df_daily_notes = pd.concat([df_daily_notes, pd.DataFrame([new_row])], ignore_index=True)
    return df_daily_notes


def _is_blank_value(v):
    """判斷一個值是否算「空」：NaN、None，或去除空白後的空字串、或 False。"""
    if pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    if isinstance(v, bool) and v is False:
        return True
    return False


def _row_has_meaningful_override(row):
    """
    判斷一列是否還有「值得保留」的手動設定：
    內控預計車次有填真的數字、備註有寫字、或計入工期被設成True，這三個才算數。
    差異/當日運棄量/累計運棄量/剩餘土方量/實際車次這些「事實」欄位不列入判斷，
    因為它們本來就只是同步時的快照，不影響任何計算（計算永遠會重新從派車紀錄算），
    舊快照留著沒有意義，可以放心清掉。
    """
    if not _is_blank_value(row.get('內控預計車次')):
        return True
    if not _is_blank_value(row.get('備註')):
        return True
    if row.get('計入工期') is True:
        return True
    return False


def _clear_daily_notes_field_or_drop(df_daily_notes, stage_choice, date_str, field):
    """
    把某一天某個欄位清空（設回 NaN，代表恢復系統預設行為）。
    清空後如果這一列已經沒有任何值得保留的手動設定，就把整列刪掉，避免留下空列造成雲端資料肥大、雜亂。
    """
    df_daily_notes = df_daily_notes.copy()
    mask = (df_daily_notes['階段名稱'] == stage_choice) & (df_daily_notes['日期'].astype(str) == str(date_str))
    if not mask.any():
        return df_daily_notes
    df_daily_notes.loc[mask, field] = np.nan
    row = df_daily_notes.loc[mask].iloc[0]
    if not _row_has_meaningful_override(row):
        df_daily_notes = df_daily_notes[~mask]
    return df_daily_notes


def _save_daily_notes_sorted(df_daily_notes):
    """存檔前依「階段名稱、日期」排序，避免雲端表格日期跳來跳去、難以閱讀核對。"""
    if not df_daily_notes.empty and '日期' in df_daily_notes.columns:
        sort_cols = ['階段名稱', '日期'] if '階段名稱' in df_daily_notes.columns else ['日期']
        df_daily_notes = df_daily_notes.sort_values(sort_cols).reset_index(drop=True)
    return save_sheet_data("stage_daily_notes", df_daily_notes)


def sync_stage_daily_log(stage_choice, df_results):
    """
    把「目前作業階段」今天這一列的統計數字，即時 upsert 寫入 stage_daily_notes 雲端分頁。
    設計給「批量設定出土分區」按下去之後自動呼叫，不需要使用者再手動去階段管控頁按儲存。

    注意：只同步「事實」欄位（實際車次/差異/當日運棄量/累計運棄量/剩餘土方量，這些是從派車
    紀錄算出來的真實數字）。「內控預計車次」跟「備註」是使用者的排程規劃值，這裡刻意不覆寫，
    只能透過階段管控頁的「💾 儲存每日車次目標與備註」手動編輯儲存。這樣做是為了讓「內控預計
    車次」在使用者調整上面的階段參數（開始/結束日期、預估土方量等）時，只要那天沒被手動改過，
    包含過去日期在內都能持續套用最新算出來的預設值，而不會被自動同步寫入的舊數字卡住。
    """
    overview = compute_stage_overview(stage_choice, df_results)
    df_range = overview["df_range"]
    df_daily_notes = overview["df_daily_notes"]

    today_str = overview["today"].strftime("%Y-%m-%d")
    today_rows = df_range[df_range['日期'] == today_str]
    if today_rows.empty:
        return False

    row = today_rows.iloc[0].to_dict()
    fact_cols = ['實際車次', '差異', '當日運棄量', '累計運棄量', '剩餘土方量']

    mask = (df_daily_notes['階段名稱'] == stage_choice) & (df_daily_notes['日期'].astype(str) == today_str)
    if mask.any():
        # 該日已有既存紀錄（可能含使用者手動設定的內控預計車次/備註），只更新事實欄位，不動那兩欄
        for col in fact_cols:
            df_daily_notes.loc[mask, col] = row.get(col, "")
        final_notes = df_daily_notes
    else:
        # 該日還沒有任何紀錄，新增一筆；內控預計車次留空，讓它之後持續套用系統預設值，不鎖住
        new_row = {col: row.get(col, "") for col in fact_cols}
        new_row['階段名稱'] = stage_choice
        new_row['日期'] = today_str
        new_row['內控預計車次'] = np.nan
        new_row['備註'] = ""
        new_row_df = pd.DataFrame([new_row])
        final_notes = pd.concat([df_daily_notes, new_row_df], ignore_index=True)

    return _save_daily_notes_sorted(final_notes)


def generate_backend_map(df_results, zone_grouped, stage_choice=None):
    """
    產出後端地圖圖片（給PDF報表嵌入用）。
    stage_choice 若給定且為第1~4挖，會跟畫面上的單階段地圖一致：
      - 第3、4挖自動排除滯洪池分區（滯洪池只有到第2挖）
      - 套用三色進度（尚未開始/進行中/已完成），只反映該階段自己的進度
    stage_choice 為 None 或「開挖前土方」時，維持原本總體4階累計完成度上色。
    """
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    from matplotlib.font_manager import FontProperties
    import gc

    fig, ax = plt.subplots(figsize=(10, 6))
    font_path = "font.ttf"
    my_font = FontProperties(fname=font_path) if os.path.exists(font_path) else None

    vol_dict = {}
    if not zone_grouped.empty:
        vol_dict = zone_grouped.set_index('出土分區')['累計實挖方量'].to_dict()
    stage_dict = df_results.set_index('分區代號')['各階累計方量'].to_dict() if not df_results.empty else {}

    stage_idx = get_stage_index(stage_choice) if stage_choice else None

    plot_df = df_results
    if stage_idx in (2, 3):
        plot_df = df_results[~df_results['分區代號'].str.contains("滯")]

    for idx, row in plot_df.iterrows():
        grid_id = row['分區代號']
        current_vol = vol_dict.get(grid_id, 0)
        thresholds = stage_dict.get(grid_id, [])
        fill_color = '#F0F0F0'

        if stage_idx is not None and thresholds and stage_idx < len(thresholds):
            band_start = thresholds[stage_idx - 1] if stage_idx > 0 else 0
            band_end = thresholds[stage_idx]
            band_size = band_end - band_start
            pct = min(max((current_vol - band_start) / band_size * 100, 0), 100) if band_size > 0 else 100.0
            if pct >= 98:
                fill_color = '#2ECC71'
            elif pct > 0:
                fill_color = '#E67E22'
            else:
                fill_color = '#F0F0F0'
        elif stage_idx is None:
            if pd.notnull(current_vol) and current_vol > 0 and len(thresholds) > 0:
                if current_vol >= thresholds[-1] * 0.98:
                    fill_color = '#2ECC71'
                else:
                    colors = ['#F1C40F', '#E67E22', '#3498DB', '#9B59B6']
                    for s_idx, t_vol in enumerate(thresholds):
                        if current_vol < t_vol * 0.98:
                            fill_color = colors[s_idx] if s_idx < len(colors) else colors[-1]
                            break

        xy = [[row['x_min'], row['y_min']], [row['x_max'], row['y_min']],
              [row['x_max'], row['y_max']], [row['x_min'], row['y_max']]]
        poly = patches.Polygon(xy, closed=True, facecolor=fill_color, edgecolor='gray', alpha=0.8)
        ax.add_patch(poly)

        if my_font:
            ax.text(row['x_center'], row['y_center'], grid_id, ha='center', va='center', fontsize=8, color='black', fontproperties=my_font)
        else:
            ax.text(row['x_center'], row['y_center'], grid_id, ha='center', va='center', fontsize=8, color='black')

    ax.autoscale_view()
    ax.set_aspect('equal')
    plt.axis('off')

    tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    plt.savefig(tmp_img.name, bbox_inches='tight', dpi=150)
    fig.clf()
    plt.close('all')
    gc.collect()
    return tmp_img.name


# ============================================================================
# 新增：每日出土 PDF 報表產出（原本遺失的功能）
# ============================================================================
def generate_daily_report_pdf(report_text, breakdown_text, display_df, map_img_path,
                               period_label, start_date, end_date, stage_overview=None,
                               daily_control_df=None, map_legend_items=None, stage_label=None):
    """
    產出每日/區間出土統計 PDF 報表：
      - 標題（置中）
      - 左欄：階段總覽（stage_overview，對應圖1紫色表格；若為 None 則退回顯示 report_text 本日回報文字）
      - 右欄：聯單分類出土明細 (breakdown_text)
      - 圖例（真正的顏色色塊，map_legend_items = [(label, hex_color), ...]） + 各區開挖階段狀態地圖
      - 各分區挖掘進度總表（放在地圖後面，避免分區數增加時把地圖往下擠）
      - 每日出土管控明細 (daily_control_df, 放在PDF最後面；標題用 stage_label)
    回傳暫存 PDF 檔案路徑。
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.colors import HexColor
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    font_path = "font.ttf"
    font_name = "CustomFont"
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont(font_name, font_path))
        except Exception:
            font_name = "Helvetica"
    else:
        font_name = "Helvetica"

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp_pdf.name, pagesize=A4)
    width, height = A4
    margin = 18 * mm
    y = height - margin

    def new_page():
        nonlocal y
        c.showPage()
        c.setFont(font_name, 10)
        y = height - margin

    # 標題（置中）
    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, y, f"CDC土方開挖{period_label}回報")
    y -= 12 * mm

    # 左欄：階段總覽（若有帶入 stage_overview，取代原本的本日回報文字）；右欄：聯單分類出土
    col2_x_top = margin + (width - 2 * margin) * 0.58
    y_left_start = y
    y_right_start = y

    y_left = y_left_start
    if stage_overview is not None:
        so = stage_overview
        c.setFont(font_name, 12)
        c.drawString(margin, y_left, f"【{so['stage_choice']}】階段總覽")
        y_left -= 7 * mm
        c.setFont(font_name, 9.5)
        overview_lines = [
            f"今天日期：{so['today']}",
            f"目前作業工期：{so['current_work_days']} 天",
            f"預計開始時間：{so['est_start']}　實際開始時間：{so['actual_start_date']}",
            f"預計施作工期：{so['est_days']} 天　推估剩餘天數：{so['remain_days']} 天",
            f"預計完成日期：{so['est_end']}",
            f"推估完成日期：{so['est_completion_date']}　（差異：{so['diff_days']} 天）",
            f"預估土方量(鬆方)：{so['est_vol']:,.1f} m³",
            f"本日出土量：{so['today_vol']:,.1f} m³",
            f"預估每日出土量：{so['est_daily_vol']:,.1f} m³",
            f"累積出土數量：{so['cum_vol']:,.1f} m³",
            f"平均出土功率：{so['avg_vol_per_day']:,.1f} m³/天",
            f"剩餘土方量：{so['remain_vol']:,.1f} m³",
            f"完成百分比：{so['percent_done']}%",
        ]
        for line in overview_lines:
            if y_left < margin + 10 * mm:
                break
            c.drawString(margin, y_left, line)
            y_left -= 5.5 * mm
    else:
        # 未帶入 stage_overview 時，退回顯示原本的本日回報文字（向下相容）
        c.setFont(font_name, 10)
        for line in str(report_text).split("\n"):
            if y_left < margin + 10 * mm:
                break
            c.drawString(margin, y_left, line)
            y_left -= 5.5 * mm

    c.setFont(font_name, 11)
    y_right = y_right_start
    c.drawString(col2_x_top, y_right, "聯單分類出土：")
    y_right -= 6 * mm
    c.setFont(font_name, 10)
    for line in str(breakdown_text).split("\n"):
        if line.strip() == "":
            continue
        c.drawString(col2_x_top, y_right, line)
        y_right -= 5.5 * mm

    y = min(y_left, y_right) - 4 * mm
    if y < margin + 20 * mm:
        new_page()

    def draw_legend(legend_items, start_y):
        """畫真正的顏色色塊圖例（不是emoji文字，避免字型不支援顯示成方框），超出頁寬會自動換行。"""
        nonlocal y
        box_size = 3.5 * mm
        gap = 5 * mm
        line_h = 6 * mm
        x_pos = margin
        cur_y = start_y
        c.setFont(font_name, 9)
        for label, hex_color in legend_items:
            text_w = c.stringWidth(label, font_name, 9)
            item_w = box_size + 1.5 * mm + text_w
            if x_pos + item_w > width - margin and x_pos > margin:
                x_pos = margin
                cur_y -= line_h
            c.setFillColor(HexColor(hex_color))
            c.rect(x_pos, cur_y - box_size + 1, box_size, box_size, fill=1, stroke=0)
            c.setFillColor(HexColor("#000000"))
            c.drawString(x_pos + box_size + 1.5 * mm, cur_y - box_size + 1, label)
            x_pos += item_w + gap
        y = cur_y - box_size - 4 * mm

    # 各區開挖階段狀態地圖（放在分區表格前面，避免分區數日後增加時把地圖一直往下擠）
    if map_img_path and os.path.exists(map_img_path):
        y -= 4 * mm
        if y < margin + 95 * mm:
            new_page()
        c.setFont(font_name, 11)
        c.drawString(margin, y, "各區開挖階段狀態圖：")
        y -= 6 * mm
        legend_items = map_legend_items if map_legend_items else [
            ("尚未開始", "#F0F0F0"), ("進行中", "#E67E22"), ("已完成", "#2ECC71"),
        ]
        draw_legend(legend_items, y)
        img_w = width - 2 * margin
        img_h = img_w * 0.6
        if y - img_h < margin:
            new_page()
        c.drawImage(map_img_path, margin, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, mask='auto')
        y -= (img_h + 6 * mm)

    # 各分區進度表格
    if y < margin + 30 * mm:
        new_page()
    c.setFont(font_name, 11)
    c.drawString(margin, y, "各分區挖掘進度總表：")
    y -= 7 * mm

    if display_df is not None and not display_df.empty:
        col_names = list(display_df.columns)
        col_width = (width - 2 * margin) / max(len(col_names), 1)
        c.setFont(font_name, 9)
        for i, col in enumerate(col_names):
            c.drawString(margin + i * col_width, y, str(col))
        y -= 2 * mm
        c.line(margin, y, width - margin, y)
        y -= 5 * mm

        for _, row in display_df.iterrows():
            if y < margin + 12 * mm:
                new_page()
                c.setFont(font_name, 9)
                for i, col in enumerate(col_names):
                    c.drawString(margin + i * col_width, y, str(col))
                y -= 2 * mm
                c.line(margin, y, width - margin, y)
                y -= 5 * mm
            for i, col in enumerate(col_names):
                c.drawString(margin + i * col_width, y, str(row[col]))
            y -= 5.5 * mm
    else:
        c.setFont(font_name, 10)
        c.drawString(margin, y, "（尚無分區資料）")
        y -= 6 * mm

    # 每日出土管控明細（放在PDF最後面，根據目前選擇的階段從雲端資料顯示）
    if daily_control_df is not None and not daily_control_df.empty:
        new_page()
        c.setFont(font_name, 13)
        stage_label_for_table = stage_label or (stage_overview['stage_choice'] if stage_overview else "")
        c.drawString(margin, y, f"【{stage_label_for_table}】每日出土管控明細")
        y -= 9 * mm

        dc_cols = list(daily_control_df.columns)
        # 欄位寬度：日期與備註給寬一點，其餘平分
        wide_cols = {'日期', '備註'}
        n_wide = sum(1 for col in dc_cols if col in wide_cols)
        n_narrow = len(dc_cols) - n_wide
        total_w = width - 2 * margin
        wide_w = total_w * 0.20
        narrow_w = (total_w - wide_w * n_wide) / max(n_narrow, 1)
        col_widths = [wide_w if col in wide_cols else narrow_w for col in dc_cols]

        def draw_dc_header():
            nonlocal y
            c.setFont(font_name, 8)
            x_pos = margin
            for col, w in zip(dc_cols, col_widths):
                c.drawString(x_pos, y, str(col))
                x_pos += w
            y -= 2 * mm
            c.line(margin, y, width - margin, y)
            y -= 4.5 * mm

        red_cols = {'差異', '累積差異'}
        draw_dc_header()
        c.setFont(font_name, 8)
        for _, row in daily_control_df.iterrows():
            if y < margin + 10 * mm:
                new_page()
                c.setFont(font_name, 13)
                c.drawString(margin, y, f"【{stage_label_for_table}】每日出土管控明細（續）")
                y -= 9 * mm
                draw_dc_header()
            x_pos = margin
            for col, w in zip(dc_cols, col_widths):
                val = row[col]
                if isinstance(val, float):
                    text_val = f"{val:,.1f}"
                else:
                    text_val = str(val)
                try:
                    is_negative = col in red_cols and float(val) < 0
                except (ValueError, TypeError):
                    is_negative = False
                if is_negative:
                    c.setFillColor(HexColor("#e74c3c"))
                c.drawString(x_pos, y, text_val)
                if is_negative:
                    c.setFillColor(HexColor("#000000"))
                x_pos += w
            y -= 5 * mm

    c.save()
    return tmp_pdf.name


# ============================================================================
# 補充：聯單交付簽收 PDF 報表（tab_delivery 原本也缺少此函式定義）
# ============================================================================
def generate_delivery_pdf(df, scope):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import base64
    import io

    font_path = "font.ttf"
    font_name = "CustomFont"
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont(font_name, font_path))
        except Exception:
            font_name = "Helvetica"
    else:
        font_name = "Helvetica"

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(tmp_pdf.name, pagesize=A4)
    width, height = A4
    margin = 18 * mm

    c.setFont(font_name, 16)
    c.drawString(margin, height - margin, f"CDC土方聯單交付簽收報表 - {scope}")
    c.setFont(font_name, 9)
    c.drawString(margin, height - margin - 7 * mm, f"匯出時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    y = height - margin - 20 * mm
    row_block_h = 55 * mm

    for _, row in df.iterrows():
        if y < row_block_h:
            c.showPage()
            c.setFont(font_name, 9)
            y = height - margin

        c.setFont(font_name, 10)
        text_lines = [
            f"交付日期時間: {row.get('交付日期','')} {row.get('交付時間','')}",
            f"廠商名稱: {row.get('廠商名稱','')}",
            f"聯單類型: {row.get('聯單類型','')}",
            f"起始序號: {row.get('起始序號','')}",
            f"發放張數: {row.get('發放張數','')}",
            f"簽收人姓名: {row.get('簽收人姓名','')}",
        ]
        ty = y
        for line in text_lines:
            c.drawString(margin, ty, line)
            ty -= 5 * mm

        sign_data = row.get('簽名資料', '')
        if isinstance(sign_data, str) and sign_data:
            try:
                img_bytes = base64.b64decode(sign_data)
                tmp_sign = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                with open(tmp_sign.name, "wb") as sf:
                    sf.write(img_bytes)
                c.drawImage(tmp_sign.name, 120 * mm, y - 30 * mm, width=60 * mm, height=35 * mm,
                            preserveAspectRatio=True, mask='auto')
            except Exception:
                c.drawString(120 * mm, y - 15 * mm, "(簽名載入失敗)")
        else:
            c.drawString(120 * mm, y - 15 * mm, "(無簽名資料)")

        c.line(margin, y - 38 * mm, width - margin, y - 38 * mm)
        y -= row_block_h

    c.save()
    return tmp_pdf.name


e_ext = 3.25
dx1 = [8.7, 8.7, 8.7, 8.7, 8.7, 10.2]
dy1 = [-9.6, -8.4, -7.5, -7.5, -7.5]
y_labels1 = ["A", "B", "C", "D", "E"]
dx2 = [6.9, 9.0, 9.0, 9.3, 9.3, 9.3, 9.3, 9.0, 9.0, 6.0]
dy2 = [-11.25, -9.0, -9.3, -9.3, -9.3, -7.5]
y_labels2 = ["A", "B'", "C'", "D'", "E'", "F'"]

df_results = pd.DataFrame()
try:
    base_x = base_x_input / scale_factor
    base_y = base_y_input / scale_factor
    x_coords1 = [base_x] + list(base_x + np.cumsum(dx1))
    y_coords1 = [base_y] + list(base_y + np.cumsum(dy1))
    x_offset = x_coords1[-1]
    x_coords2 = [x_offset] + list(x_offset + np.cumsum(dx2))
    y_coords2 = [base_y] + list(base_y + np.cumsum(dy2))

    depths_admin = get_thickness_from_gl(gl_admin_input, current_gl)
    depths_lab = get_thickness_from_gl(gl_lab_input, current_gl)
    depths_bc = get_thickness_from_gl(gl_bc_input, current_gl)
    depths_a = get_thickness_from_gl(gl_a_input, current_gl)

    results = []

    for j in range(len(dy1)):
        for i in range(len(dx1)):
            if j >= 2 and i >= 3: continue
            grid_id = f"{y_labels1[j]}{i+1}"
            x_min, x_max = x_coords1[i], x_coords1[i+1]
            y_max, y_min = y_coords1[j], y_coords1[j+1]
            if grid_id in ["E1", "E2", "E3"]: y_min -= e_ext
            poly = Polygon([(x_min, y_min), (x_max, y_min), (x_max, y_max), (x_min, y_max)])
            vols = [poly.area * d for d in depths_admin]
            cum_vols = [round(v, 0) for v in list(np.cumsum(vols))]
            v1 = vols[0] if len(vols) > 0 else 0
            v2 = vols[1] if len(vols) > 1 else 0
            v3 = vols[2] if len(vols) > 2 else 0
            v4 = vols[3] if len(vols) > 3 else 0
            results.append({
                "分區代號": grid_id, "區域面積(㎡)": round(poly.area, 0),
                "第1挖方量(m³)": round(v1, 0), "第2挖方量(m³)": round(v2, 0),
                "第3挖方量(m³)": round(v3, 0), "第4挖方量(m³)": round(v4, 0),
                "預估總土方": round(sum(vols), 0), "各階累計方量": cum_vols,
                "x_min": x_min, "x_max": x_max, "y_min": y_min, "y_max": y_max,
                "x_center": (x_min + x_max)/2, "y_center": (y_min + y_max)/2
            })

    for j in range(len(dy2)):
        for i in range(len(dx2)):
            grid_id = f"{y_labels2[j]}{i+7}"
            x_min, x_max = x_coords2[i], x_coords2[i+1]
            y_max, y_min = y_coords2[j], y_coords2[j+1]
            poly = Polygon([(x_min, y_min), (x_max, y_min), (x_max, y_max), (x_min, y_max)])
            vols = [poly.area * d for d in depths_lab]
            cum_vols = [round(v, 0) for v in list(np.cumsum(vols))]
            v1 = vols[0] if len(vols) > 0 else 0
            v2 = vols[1] if len(vols) > 1 else 0
            v3 = vols[2] if len(vols) > 2 else 0
            v4 = vols[3] if len(vols) > 3 else 0
            results.append({
                "分區代號": grid_id, "區域面積(㎡)": round(poly.area, 0),
                "第1挖方量(m³)": round(v1, 0), "第2挖方量(m³)": round(v2, 0),
                "第3挖方量(m³)": round(v3, 0), "第4挖方量(m³)": round(v4, 0),
                "預估總土方": round(sum(vols), 0), "各階累計方量": cum_vols,
                "x_min": x_min, "x_max": x_max, "y_min": y_min, "y_max": y_max,
                "x_center": (x_min + x_max)/2, "y_center": (y_min + y_max)/2
            })

    bc_x = [-2764.56, -2758.41, -2749.46]
    bc_y = [-250.94, -256.69, -262.94, -270.04, -275.14]
    idx_l = 1
    for i in range(len(bc_x)-1):
        for j in range(len(bc_y)-1):
            old_idx = j * 2 + i + 1
            if old_idx in [1, 3]: continue
            grid_id = f"滯BC{idx_l}"
            x_min, x_max = bc_x[i], bc_x[i+1]
            y_max, y_min = bc_y[j], bc_y[j+1]
            poly = Polygon([(x_min, y_min), (x_max, y_min), (x_max, y_max), (x_min, y_max)])
            vols = [poly.area * d for d in depths_bc]
            cum_vols = [round(v, 0) for v in list(np.cumsum(vols))]
            v1 = vols[0] if len(vols) > 0 else 0
            v2 = vols[1] if len(vols) > 1 else 0
            v3 = vols[2] if len(vols) > 2 else 0
            v4 = vols[3] if len(vols) > 3 else 0
            results.append({
                "分區代號": grid_id, "區域面積(㎡)": round(poly.area, 0),
                "第1挖方量(m³)": round(v1, 0), "第2挖方量(m³)": round(v2, 0),
                "第3挖方量(m³)": round(v3, 0), "第4挖方量(m³)": round(v4, 0),
                "預估總土方": round(sum(vols), 0), "各階累計方量": cum_vols,
                "x_min": x_min, "x_max": x_max, "y_min": y_min, "y_max": y_max,
                "x_center": (x_min + x_max)/2, "y_center": (y_min + y_max)/2
            })
            idx_l += 1

    a_x = [-2606.06, -2592.82]
    a_y = [-276.14, -284.44, -290.24, -296.04]
    idx_r = 1
    for j in range(len(a_y)-1):
        for i in range(len(a_x)-1):
            x_min, x_max = a_x[i], a_x[i+1]
            y_max, y_min = a_y[j], a_y[j+1]
            grid_id = f"滯洪池A{idx_r}"
            poly = Polygon([(x_min, y_min), (x_max, y_min), (x_max, y_max), (x_min, y_max)])
            vols = [poly.area * d for d in depths_a]
            cum_vols = [round(v, 0) for v in list(np.cumsum(vols))]
            v1 = vols[0] if len(vols) > 0 else 0
            v2 = vols[1] if len(vols) > 1 else 0
            v3 = vols[2] if len(vols) > 2 else 0
            v4 = vols[3] if len(vols) > 3 else 0
            results.append({
                "分區代號": grid_id, "區域面積(㎡)": round(poly.area, 0),
                "第1挖方量(m³)": round(v1, 0), "第2挖方量(m³)": round(v2, 0),
                "第3挖方量(m³)": round(v3, 0), "第4挖方量(m³)": round(v4, 0),
                "預估總土方": round(sum(vols), 0), "各階累計方量": cum_vols,
                "x_min": x_min, "x_max": x_max, "y_min": y_min, "y_max": y_max,
                "x_center": (x_min + x_max)/2, "y_center": (y_min + y_max)/2
            })
            idx_r += 1

    df_results = pd.DataFrame(results)
except Exception as e:
    st.sidebar.error(f"圖資運算錯誤: {e}")

tab_grid, tab_stats, tab_stage, tab_sync, tab_manifest, tab_delivery = st.tabs([
    "🗺️ 圖資與方量基準", "📊 出土統計儀表板", "📈 階段開挖管控", "🧾 官方聯單對帳", "🎫 聯單庫存管理", "✍️ 現場廠商簽收"
])

with tab_grid:
    export_columns = ['分區代號', '區域面積(㎡)', '第1挖方量(m³)', '第2挖方量(m³)', '第3挖方量(m³)', '第4挖方量(m³)', '預估總土方']
    if st.button("🚀 推送分區資料至雲端試算表"):
        if save_sheet_data("grid_zones", df_results[export_columns]):
            st.success("分區基準已成功上傳！")

    col1, col2 = st.columns([3, 2])
    with col2:
        st.write("### 基準方量總表")
        st.dataframe(df_results[export_columns], height=600)
        st.success(f"全區預估總土方量： **{df_results['預估總土方'].sum():,.0f} m³**")

    with col1:
        st.write("### 精準網格地圖")
        fig = go.Figure()
        for idx, row in df_results.iterrows():
            fig.add_trace(go.Scatter(
                x=[row['x_min'], row['x_max'], row['x_max'], row['x_min'], row['x_min']],
                y=[row['y_min'], row['y_min'], row['y_max'], row['y_max'], row['y_min']],
                mode='lines', line=dict(color='blue', width=1),
                fill='toself', fillcolor='rgba(0, 100, 255, 0.1)', showlegend=False, hoverinfo='skip'
            ))
            fig.add_annotation(x=row['x_center'], y=row['y_center'], text=row['分區代號'], showarrow=False, font=dict(color="red", size=12))
        fig.update_layout(dragmode='pan', xaxis_title="X (m)", yaxis_title="Y (m)", yaxis=dict(scaleanchor="x", scaleratio=1), height=700, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': True, 'displayModeBar': False})

with tab_stats:
    st.write("### 📊 雲端出土統計儀表板")

    df_logs = load_sheet_data("dispatch_logs")
    if not df_logs.empty:
        df_logs['orig_index'] = df_logs.index

    tw_today = (datetime.utcnow() + timedelta(hours=8)).date()
    st.markdown("#### 📅 篩選統計時間區間")
    date_selection = st.date_input("選擇區間 (單選一日或拖曳選擇範圍)：", value=(tw_today, tw_today))

    if isinstance(date_selection, tuple) and len(date_selection) == 2:
        start_date, end_date = date_selection
    elif isinstance(date_selection, tuple) and len(date_selection) == 1:
        start_date = end_date = date_selection[0]
    else:
        start_date = end_date = date_selection

    delta_days = (end_date - start_date).days
    if delta_days == 0:
        period_label = "本日"
    elif start_date.weekday() == 0 and end_date.weekday() == 6 and delta_days == 6:
        period_label = "本週"
    elif start_date.day == 1 and (end_date + timedelta(days=1)).day == 1 and start_date.month == end_date.month:
        period_label = "本月"
    else:
        period_label = "本區間"

    # PDF 產出用暫存變數
    pdf_report_text = ""
    pdf_breakdown_text = ""
    pdf_display_df = pd.DataFrame()

    if not df_logs.empty and "日期" in df_logs.columns:
        if "聯單序號" not in df_logs.columns:
            df_logs["聯單序號"] = ""

        df_logs['ParsedDate'] = pd.to_datetime(df_logs['日期']).dt.date
        valid_logs = df_logs.copy()

        if not valid_logs.empty and '時間' in valid_logs.columns:
            valid_logs = valid_logs.sort_values(['車頭車號', '日期', '時間'])

        range_logs = valid_logs[(valid_logs['ParsedDate'] >= start_date) & (valid_logs['ParsedDate'] <= end_date)]
        cumul_logs = valid_logs[valid_logs['ParsedDate'] <= end_date].copy()

        range_trucks = range_logs['車頭車號'].nunique() if '車頭車號' in range_logs.columns else 0
        range_trips = len(range_logs)
        range_vol = pd.to_numeric(range_logs['載運方量(m³)'], errors='coerce').sum() if '載運方量(m³)' in range_logs.columns else 0
        total_all_trips = len(cumul_logs)

        period_days = range_logs['ParsedDate'].nunique() if not range_logs.empty and 'ParsedDate' in range_logs.columns else 0
        period_rate = round(range_trips / period_days, 1) if period_days > 0 else 0

        # 累計總天數：改用日曆天數（從第一筆出土紀錄那天算到統計結束日，含中間沒出土的日期），
        # 跟「階段總覽」的「目前作業工期」算法一致，不再只算「有出土紀錄」的天數
        if not cumul_logs.empty and 'ParsedDate' in cumul_logs.columns:
            earliest_log_date = cumul_logs['ParsedDate'].min()
            total_days = (end_date - earliest_log_date).days + 1
        else:
            total_days = 0
        total_rate = round(total_all_trips / total_days, 1) if total_days > 0 else 0

        st.markdown(f"#### 📊 {period_label}統計結果 ({start_date} 至 {end_date})")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(f"{period_label}出車車頭數", f"{range_trucks} 輛")
        m2.metric(f"{period_label}總車次", f"{range_trips} 台")
        m3.metric(f"{period_label}實挖方量", f"{range_vol:,.0f} m³")
        m4.metric(f"{period_label}出土功率", f"{period_rate} 台/天")
        m5.metric("總出土功率", f"{total_rate} 台/天")
        st.divider()

        zone_grouped = pd.DataFrame()

        display_df = pd.DataFrame()
        if '出土分區' in cumul_logs.columns and '載運方量(m³)' in cumul_logs.columns:
            df_assigned = cumul_logs[cumul_logs['出土分區'] != '未指定'].copy()
            if not df_assigned.empty:
                df_assigned['載運方量(m³)'] = pd.to_numeric(df_assigned['載運方量(m³)'], errors='coerce')
                zone_grouped = df_assigned.groupby('出土分區')['載運方量(m³)'].sum().reset_index()
                zone_grouped.rename(columns={'載運方量(m³)': '累計實挖方量'}, inplace=True)

                baseline_dict = df_results.set_index('分區代號')['預估總土方'].to_dict() if not df_results.empty else {}
                zone_grouped['預估基準方量'] = zone_grouped['出土分區'].map(baseline_dict)
                zone_grouped['完成率數值'] = (zone_grouped['累計實挖方量'] / zone_grouped['預估基準方量'] * 100).round(1)

                zone_grouped['累計實挖方量_顯示'] = zone_grouped['累計實挖方量'].apply(lambda x: f"{x:,.0f}")
                zone_grouped['預估基準方量_顯示'] = zone_grouped['預估基準方量'].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else '無基準量')
                zone_grouped['完成率_顯示'] = zone_grouped['完成率數值'].fillna('不適用').astype(str)
                zone_grouped['完成率_顯示'] = zone_grouped['完成率_顯示'].apply(lambda x: f"{x}%" if x != '不適用' else x)

                display_df = zone_grouped[['出土分區', '累計實挖方量_顯示', '預估基準方量_顯示', '完成率_顯示']].rename(
                    columns={'累計實挖方量_顯示': '累計實挖方量', '預估基準方量_顯示': '預估基準方量', '完成率_顯示': '完成率(%)'}
                )

        st.markdown(f"#### 📱 {period_label}回報")

        total_excavated = zone_grouped[zone_grouped['出土分區'] != '開挖前土方']['累計實挖方量'].sum() if not zone_grouped.empty else 0
        pre_excavated = zone_grouped[zone_grouped['出土分區'] == '開挖前土方']['累計實挖方量'].sum() if not zone_grouped.empty and '開挖前土方' in zone_grouped['出土分區'].values else 0

        manifest_breakdown_str = ""
        if '聯單序號' in range_logs.columns and '載運方量(m³)' in range_logs.columns:
            range_serials = range_logs['聯單序號'].fillna('').astype(str).str.strip().str.upper()
            range_vols = pd.to_numeric(range_logs['載運方量(m³)'], errors='coerce').fillna(0)

            m_types = ["B1", "B2-3", "B4", "B5"]
            breakdown_lines = []
            for m_type in m_types:
                mask = range_serials.str.contains(m_type.upper(), regex=False)
                m_trips = mask.sum()
                m_vol = range_vols[mask].sum()
                breakdown_lines.append(f"* {m_type} 聯單： {m_trips} 台 / {m_vol:,.0f} m³")
            manifest_breakdown_str = "\n".join(breakdown_lines)

        manifest_total = 79692.0
        overall_rate = round((total_excavated / manifest_total * 100), 1) if manifest_total > 0 else 0

        report_text_left = f"""【CDC土方開挖{period_label}回報】 區間: {start_date} 至 {end_date}
{period_label}出土天數： {period_days} 天
{period_label}車次： {range_trips} 台
{period_label}出土功率： {period_rate} 台/天
{period_label}出土方量： {range_vol:,.0f} m³
累計總天數： {total_days} 天
累計總車次： {total_all_trips} 台
總出土功率： {total_rate} 台/天
累計實挖方量： {total_excavated:,.0f} m³ (另計開挖前土方: {pre_excavated:,.0f} m³)
聯單預估總出土： {manifest_total:,.0f} m³
總體開挖進度： {overall_rate}%"""

        report_text_right = f"區間聯單分類出土：\n{manifest_breakdown_str}"
        ui_display_text = f"{report_text_left}\n\n{report_text_right}"

        # 保留給 PDF 匯出使用
        pdf_report_text = report_text_left
        pdf_breakdown_text = manifest_breakdown_str
        pdf_display_df = display_df

        col_txt, col_fig = st.columns([1, 2])
        with col_txt:
            st.info(ui_display_text.replace("\n", "\n\n"))

        with col_fig:
            _stats_stage_idx = get_stage_index(global_stage_choice)
            if _stats_stage_idx is not None:
                st.markdown(f"**進度圖例說明：（依目前作業階段【{global_stage_choice}】上色）**")
                st.markdown("⬜ 尚未開始 🟧 進行中 🟩 已完成")
            else:
                st.markdown("**進度圖例說明：（開挖前土方無分區門檻，顯示總體4階累計完成度）**")
                st.markdown("⬜ 尚未開挖 🟨 1挖進行中 🟧 1挖完成/2挖進行中 🟦 2挖完成/3挖進行中 🟪 3挖完成/4挖進行中 🟩 開挖完成")
            fig_map = go.Figure()
            if not df_results.empty:
                vol_dict = {}
                if not zone_grouped.empty:
                    vol_dict = zone_grouped.set_index('出土分區')['累計實挖方量'].to_dict()

                stage_dict = df_results.set_index('分區代號')['各階累計方量'].to_dict()

                # 第3、4挖沒有滯洪池分區（滯洪池只有2挖），地圖上要把這些格子濾掉
                if _stats_stage_idx in (2, 3):
                    df_map_zones = df_results[~df_results['分區代號'].str.contains("滯")]
                else:
                    df_map_zones = df_results

                for idx, row in df_map_zones.iterrows():
                    grid_id = row['分區代號']
                    current_vol = vol_dict.get(grid_id, 0)
                    thresholds = stage_dict.get(grid_id, [])

                    if _stats_stage_idx is not None and thresholds and _stats_stage_idx < len(thresholds):
                        band_start = thresholds[_stats_stage_idx - 1] if _stats_stage_idx > 0 else 0
                        band_end = thresholds[_stats_stage_idx]
                        band_size = band_end - band_start
                        pct = min(max((current_vol - band_start) / band_size * 100, 0), 100) if band_size > 0 else 100.0

                        if pct >= 98:
                            fill_color = 'rgba(46, 204, 113, 0.8)'
                            stage_text = f"{global_stage_choice} 已完成"
                        elif pct > 0:
                            fill_color = 'rgba(230, 126, 34, 0.7)'
                            stage_text = f"{global_stage_choice} 進行中: {pct:.0f}%"
                        else:
                            fill_color = 'rgba(240, 240, 240, 0.5)'
                            stage_text = f"{global_stage_choice} 尚未開始"
                    else:
                        stage_text = "尚未開挖"
                        fill_color = 'rgba(240, 240, 240, 0.5)'

                        if pd.notnull(current_vol) and current_vol > 0 and len(thresholds) > 0:
                            if current_vol >= thresholds[-1] * 0.98:
                                stage_text = "開挖完成"
                                fill_color = 'rgba(46, 204, 113, 0.8)'
                            else:
                                colors = ['rgba(241, 196, 15, 0.7)', 'rgba(230, 126, 34, 0.7)', 'rgba(52, 152, 219, 0.7)', 'rgba(155, 89, 182, 0.7)']
                                for s_idx, t_vol in enumerate(thresholds):
                                    if current_vol < t_vol * 0.98:
                                        if s_idx == 0:
                                            stage_text = "1挖進行中"
                                        else:
                                            stage_text = f"{s_idx}挖完成 / {s_idx+1}挖進行中"
                                        fill_color = colors[s_idx] if s_idx < len(colors) else colors[-1]
                                        break

                    fig_map.add_trace(go.Scatter(
                        x=[row['x_min'], row['x_max'], row['x_max'], row['x_min'], row['x_min']],
                        y=[row['y_min'], row['y_min'], row['y_max'], row['y_max'], row['y_min']],
                        mode='lines', line=dict(color='gray', width=1),
                        fill='toself', fillcolor=fill_color, showlegend=False, hoverinfo='text',
                        text=f"{grid_id}<br>{stage_text}<br>已挖: {current_vol:,.0f} m³"
                    ))
                    fig_map.add_annotation(x=row['x_center'], y=row['y_center'], text=grid_id, showarrow=False, font=dict(color="black", size=10))

                fig_map.update_layout(title=f"各區階數開挖狀態 (截至 {end_date})　【{global_stage_choice}】", dragmode='pan', xaxis_title="", yaxis_title="", yaxis=dict(scaleanchor="x", scaleratio=1), height=500, margin=dict(l=0, r=0, t=30, b=0))
                st.plotly_chart(fig_map, use_container_width=True, config={'displayModeBar': False})

        st.divider()

        # ====================================================================
        # 新增：每日出土 PDF 報表匯出按鍵
        # ====================================================================
        st.markdown(f"#### 📄 匯出每日出土 PDF 報表　（目前作業階段：**{global_stage_choice}**，可於左側側邊欄切換）")
        if st.button("📥 一鍵產出PDF報表", key="btn_daily_pdf", use_container_width=True):
            try:
                with st.spinner("PDF 產生中，請稍候..."):
                    map_img_path = generate_backend_map(df_results, zone_grouped, stage_choice=global_stage_choice) if not df_results.empty else None
                    stage_overview_for_pdf = compute_stage_overview(global_stage_choice, df_results)
                    daily_control_cols = ['日期', '內控預計車次', '實際車次', '差異', '累積差異', '當日運棄量', '累計運棄量', '剩餘土方量', '備註']
                    _pdf_df_range = stage_overview_for_pdf["df_range"]
                    _pdf_df_range_display = _pdf_df_range[(_pdf_df_range['當日運棄量'] > 0) | (_pdf_df_range['計入工期'] == True)]
                    daily_control_for_pdf = _pdf_df_range_display[daily_control_cols].copy()

                    if get_stage_index(global_stage_choice) is not None:
                        pdf_legend_items = [
                            ("尚未開始", "#F0F0F0"),
                            ("進行中", "#E67E22"),
                            ("已完成", "#2ECC71"),
                        ]
                    else:
                        pdf_legend_items = [
                            ("尚未開挖", "#F0F0F0"),
                            ("1挖進行中", "#F1C40F"),
                            ("1挖完成/2挖進行中", "#E67E22"),
                            ("2挖完成/3挖進行中", "#3498DB"),
                            ("3挖完成/4挖進行中", "#9B59B6"),
                            ("開挖完成", "#2ECC71"),
                        ]

                    pdf_path = generate_daily_report_pdf(
                        report_text=pdf_report_text,
                        breakdown_text=pdf_breakdown_text,
                        display_df=pdf_display_df,
                        map_img_path=map_img_path,
                        period_label=period_label,
                        start_date=start_date,
                        end_date=end_date,
                        stage_overview=None,
                        daily_control_df=daily_control_for_pdf,
                        map_legend_items=pdf_legend_items,
                        stage_label=global_stage_choice,
                    )
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label=f"✅ 點此下載 {period_label}出土報表 ({start_date}~{end_date}).pdf",
                        data=f,
                        file_name=f"daily_excavation_report_{start_date}_{end_date}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
            except Exception as ex:
                st.error(f"PDF 匯出失敗：{ex}")

        st.divider()

        st.markdown("#### ⚙️ 批量設定出土分區")
        df_unassigned = valid_logs[valid_logs['出土分區'] == '未指定'].copy()
        if not df_unassigned.empty:
            st.info(f"尚有 {len(df_unassigned)} 筆有效紀錄未指定分區，請勾選並套用。")

            select_all = st.checkbox("☑️ 一鍵全選所有未指定紀錄", value=False)
            df_unassigned.insert(0, '勾選', select_all)

            edited_unassigned = st.data_editor(df_unassigned, hide_index=True, column_config={"勾選": st.column_config.CheckboxColumn(required=True)})

            zone_list = df_results["分區代號"].tolist() if not df_results.empty else []
            col_z1, col_z2 = st.columns([2, 1])
            with col_z1:
                selected_zone = st.selectbox("選擇要套用的分區", options=["請選擇", "開挖前土方"] + zone_list)
            with col_z2:
                if st.button("套用到勾選的紀錄"):
                    if selected_zone == "請選擇":
                        st.error("請先選擇分區")
                    else:
                        checked_rows = edited_unassigned[edited_unassigned['勾選'] == True]
                        if len(checked_rows) > 0:
                            original_indices = checked_rows['orig_index'].tolist()
                            df_logs.loc[original_indices, '出土分區'] = selected_zone

                            if 'orig_index' in df_logs.columns:
                                df_logs = df_logs.drop(columns=['orig_index'])
                            if 'ParsedDate' in df_logs.columns:
                                df_logs = df_logs.drop(columns=['ParsedDate'])

                            if save_sheet_data("dispatch_logs", df_logs):
                                sync_stage_daily_log(global_stage_choice, df_results)
                                st.success(f"成功更新 {len(checked_rows)} 筆紀錄！本日「{global_stage_choice}」逐日紀錄已同步寫入雲端。")
                                st.rerun()
                        else:
                            st.warning("⚠️ 請至少勾選一筆要套用的紀錄。")
        else:
            st.success("目前所有有效紀錄皆已分配分區。")

        st.divider()
        st.markdown("#### 📍 各分區挖掘進度總表")
        if not display_df.empty:
            st.dataframe(display_df, use_container_width=True, hide_index=True)

        with st.expander("📂 檢視所有歷史紀錄"):
            clean_show_logs = df_logs.copy()
            if 'orig_index' in clean_show_logs.columns:
                clean_show_logs = clean_show_logs.drop(columns=['orig_index'])
            if 'ParsedDate' in clean_show_logs.columns:
                clean_show_logs = clean_show_logs.drop(columns=['ParsedDate'])

            show_logs = clean_show_logs.sort_values(['日期', '時間'], ascending=[False, False])
            edited_all = st.data_editor(show_logs, use_container_width=True)
            if st.button("💾 儲存歷史紀錄修改"):
                if save_sheet_data("dispatch_logs", edited_all):
                    st.success("歷史紀錄修改已儲存！")
                    st.rerun()
    else:
        st.info("尚無出土紀錄。")

with tab_stage:
    st.write("### 📈 階段開挖管控")
    st.info("此分頁用於單一開挖階段的進度追蹤。左側側邊欄切換「目前作業階段」，此頁與出土儀表板的PDF報表會同步套用該階段。")

    stage_choice = global_stage_choice
    st.markdown(f"目前檢視階段：**{stage_choice}**　（可至左側側邊欄「🎯 目前作業階段」切換）")

    df_stage_set = load_sheet_data("stage_settings")
    if df_stage_set.empty or "階段名稱" not in df_stage_set.columns:
        df_stage_set = pd.DataFrame(columns=["階段名稱", "預計開始時間", "預計結束日期", "預估土方量(鬆方)", "單車預設實方"])

    overview = compute_stage_overview(stage_choice, df_results)
    target_col = overview["target_col"]
    df_stage_map = overview["df_stage_map"]

    current_set = df_stage_set[df_stage_set["階段名稱"] == stage_choice]
    if current_set.empty:
        new_row = pd.DataFrame([{
            "階段名稱": stage_choice,
            "預計開始時間": overview["today"].strftime("%Y-%m-%d"),
            "預計結束日期": (overview["today"] + timedelta(days=20)).strftime("%Y-%m-%d"),
            "預估土方量(鬆方)": overview["est_vol_default"],
            "單車預設實方": 12.0
        }])
        display_stage_set = new_row
    else:
        display_stage_set = current_set.copy()

    display_stage_set = display_stage_set.copy()
    display_stage_set["預計開始時間"] = pd.to_datetime(display_stage_set["預計開始時間"], errors='coerce')
    display_stage_set["預計結束日期"] = pd.to_datetime(display_stage_set["預計結束日期"], errors='coerce')

    st.markdown(f"#### ⚙️ 【{stage_choice}】參數設定")
    st.caption("💡 「預計施作工期」改由系統自動算（= 預計結束日期 − 預計開始時間），你只要設定開始與結束日期即可。下方「每日出土管控明細」表格中的「內控預計車次」預設值 = 預估土方量(鬆方) ÷ 單車預設實方 ÷ 預計施作工期，所有日期會統一帶入這個算出來的數字；若某一天要單獨調整，直接在該列的「內控預計車次」欄位改掉，按下方「💾 儲存」後，那一天就會變成你手動輸入的數字，不會再被自動預設值覆蓋（其餘沒改過的日期則繼續套用預設值）。")
    edited_stage_set = st.data_editor(
        display_stage_set,
        hide_index=True,
        use_container_width=True,
        column_config={
            "預計開始時間": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "預計結束日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
        },
    )
    _preview_start = pd.to_datetime(edited_stage_set.iloc[0]["預計開始時間"]).date()
    _preview_end = pd.to_datetime(edited_stage_set.iloc[0]["預計結束日期"]).date()
    st.caption(f"📐 系統自動算出的預計施作工期：**{max((_preview_end - _preview_start).days, 1)} 天**")

    if st.button("💾 儲存本階段設定"):
        save_set = edited_stage_set.copy()
        save_set["預計開始時間"] = pd.to_datetime(save_set["預計開始時間"]).dt.strftime("%Y-%m-%d")
        save_set["預計結束日期"] = pd.to_datetime(save_set["預計結束日期"]).dt.strftime("%Y-%m-%d")
        other_sets = df_stage_set[df_stage_set["階段名稱"] != stage_choice]
        final_sets = pd.concat([other_sets, save_set], ignore_index=True)
        save_sheet_data("stage_settings", final_sets)
        st.rerun()

    # 使用畫面上剛編輯的（尚未儲存的）參數即時重算總覽，維持「編輯即時預覽」效果
    _override_row = edited_stage_set.iloc[0].copy()
    _override_row["預計開始時間"] = pd.to_datetime(_override_row["預計開始時間"]).strftime("%Y-%m-%d")
    _override_row["預計結束日期"] = pd.to_datetime(_override_row["預計結束日期"]).strftime("%Y-%m-%d")
    overview = compute_stage_overview(stage_choice, df_results, override_settings_row=_override_row)
    df_range = overview["df_range"]
    df_daily_notes = overview["df_daily_notes"]

    st.markdown(f"#### 🟣 【{stage_choice}】總覽")

    html_table = build_stage_overview_html(overview)
    st.markdown(html_table, unsafe_allow_html=True)

    st.divider()
    st.markdown("#### 📜 查詢過往報表")
    st.caption("選一個過去的日期，呈現「截至那一天」的歷史快照（該日出土量、累積量、剩餘量、完成百分比等），不影響上面即時總覽或下面的編輯表格。")
    query_col1, query_col2 = st.columns([1, 3])
    with query_col1:
        query_date = st.date_input(
            "查詢日期",
            value=overview["today"],
            min_value=overview["est_start"],
            max_value=overview["today"],
            key=f"history_query_date_{stage_choice}",
        )
    if query_date != overview["today"]:
        historical_overview = compute_stage_overview(stage_choice, df_results, as_of_date=query_date)
        st.markdown(f"##### 🟣 【{stage_choice}】截至 {query_date} 的歷史快照")
        st.markdown(build_stage_overview_html(historical_overview), unsafe_allow_html=True)
    else:
        st.caption("目前選的是今天，跟上面的即時總覽相同；選其他日期即可查詢過往快照。")

    st.divider()
    st.markdown("#### 📅 每日出土管控明細")
    st.caption("💡 「備註」欄可填寫無法出土原因（例如：台北港停止收容），供後續對帳查閱；「剩餘土方量」為該日累計後推算之剩餘量；「累積差異」為「差異」逐日累加。「差異」與「累積差異」為負數時，下方預覽表會以紅色標示。下方表格只顯示「有出土」或「被標記為特例」的日期；沒出土又沒被標記的日期不會出現、也不算進「目前作業工期」。")

    zero_vol_df = df_range[df_range['當日運棄量'] == 0][['日期', '計入工期']].copy()
    with st.expander(f"⚙️ 管理沒有出土的日期（共 {len(zero_vol_df)} 天；預設不顯示、不算工期，勾選才會顯示並算入工期）", expanded=False):
        if zero_vol_df.empty:
            st.caption("目前沒有出土量為0的日期。")
        else:
            edited_zero = st.data_editor(
                zero_vol_df.rename(columns={'計入工期': '算入工期並顯示'}),
                column_config={
                    "日期": st.column_config.TextColumn(disabled=True),
                    "算入工期並顯示": st.column_config.CheckboxColumn(default=False),
                },
                hide_index=True,
                use_container_width=True,
                key=f"zero_vol_editor_{stage_choice}",
            )
            if st.button("💾 儲存特例設定", key=f"save_zero_vol_{stage_choice}"):
                updated_notes = df_daily_notes.copy()
                for _, r in edited_zero.iterrows():
                    d = str(r['日期'])
                    if bool(r['算入工期並顯示']):
                        updated_notes = _upsert_daily_notes_field(updated_notes, stage_choice, d, '計入工期', True)
                    else:
                        # 沒勾 = 恢復預設行為，不需要特地存一筆 False，
                        # 直接清掉這個欄位；如果那天已經沒有其他值得保留的手動設定，整列一併移除，避免留下空列
                        updated_notes = _clear_daily_notes_field_or_drop(updated_notes, stage_choice, d, '計入工期')
                _save_daily_notes_sorted(updated_notes)
                st.success("已儲存特例設定！沒出土但打勾的日期會出現在下方明細表並算入工期；沒勾的日期不會在雲端留下多餘紀錄。")
                st.rerun()

    with st.expander("🧹 清理雲端舊資料（一次性；清掉沒有實際內容的空白紀錄列）", expanded=False):
        st.caption("如果你的雲端 stage_daily_notes 分頁裡有很多內控預計車次、備註都空白，計入工期也是空的舊紀錄列（例如舊版程式曾經誤寫入的資料），可以用這顆按鈕一次清掉。只會刪除「內控預計車次沒填、備註沒寫、計入工期沒被勾選為True」的列；只要三者有一個有內容，那一列就會被保留。")
        if st.button("🧹 清理本階段沒有意義的舊紀錄列", key=f"cleanup_notes_{stage_choice}"):
            cleaned_notes = df_daily_notes.copy()
            if not cleaned_notes.empty:
                mask_stage_clean = cleaned_notes['階段名稱'] == stage_choice
                keep_mask = (~mask_stage_clean) | cleaned_notes.apply(_row_has_meaningful_override, axis=1)
                removed_count = int((~keep_mask).sum())
                cleaned_notes = cleaned_notes[keep_mask]
            else:
                removed_count = 0
            _save_daily_notes_sorted(cleaned_notes)
            st.success(f"已清理【{stage_choice}】{removed_count} 筆沒有實際內容的舊紀錄列。")
            st.rerun()

    display_cols = ['日期', '內控預計車次', '實際車次', '差異', '累積差異', '當日運棄量', '累計運棄量', '剩餘土方量', '備註']
    df_range_display = df_range[(df_range['當日運棄量'] > 0) | (df_range['計入工期'] == True)]

    def _highlight_negative_diff(val):
        try:
            return 'color: #e74c3c; font-weight: bold;' if float(val) < 0 else ''
        except (ValueError, TypeError):
            return ''

    preview_styler = df_range_display[display_cols].style.applymap(_highlight_negative_diff, subset=['差異', '累積差異'])
    st.dataframe(preview_styler, use_container_width=True, hide_index=True)

    col_reset1, col_reset2 = st.columns([3, 1])
    with col_reset2:
        if st.button("🔄 重設內控預計車次為系統預設值", use_container_width=True):
            mask_stage = df_daily_notes['階段名稱'] == stage_choice
            if '內控預計車次' in df_daily_notes.columns:
                df_daily_notes.loc[mask_stage, '內控預計車次'] = np.nan
            # 清空後，把已經沒有任何值得保留設定（備註/計入工期）的空列一併移除，避免留下垃圾列
            if not df_daily_notes.empty:
                keep_mask = (~mask_stage) | df_daily_notes.apply(_row_has_meaningful_override, axis=1)
                df_daily_notes = df_daily_notes[keep_mask]
            _save_daily_notes_sorted(df_daily_notes)
            st.success(f"已清空【{stage_choice}】所有日期的內控預計車次，全部改回套用目前參數算出的系統預設值。")
            st.rerun()
    with col_reset1:
        st.caption("⚠️ 若過去曾手動調整過某幾天的內控預計車次，按下這顆按鈕會把「這個階段」所有日期都清空重算，手動設定的數字也會一併被清掉，請注意。")

    st.markdown("##### ✏️ 編輯內控預計車次 / 備註")
    edited_daily = st.data_editor(
        df_range_display[display_cols],
        column_config={
            "內控預計車次": st.column_config.NumberColumn(required=True),
            "實際車次": st.column_config.NumberColumn(disabled=True),
            "差異": st.column_config.NumberColumn(disabled=True),
            "累積差異": st.column_config.NumberColumn(disabled=True),
            "當日運棄量": st.column_config.NumberColumn(disabled=True),
            "累計運棄量": st.column_config.NumberColumn(disabled=True),
            "剩餘土方量": st.column_config.NumberColumn(disabled=True),
            "備註": st.column_config.TextColumn()
        },
        hide_index=True,
        use_container_width=True
    )

    if st.button("💾 儲存每日車次目標與備註 (全欄位寫入雲端)"):
        save_df = edited_daily.copy()
        save_df['階段名稱'] = stage_choice
        save_df['計入工期'] = True  # 這裡看得到的日期一定是「有出土」或「已標記特例」，都算工期
        # 只替換「這次編輯到的日期」，其餘（含被隱藏的0出土日期特例設定）完全不動
        edited_dates = set(save_df['日期'])
        other_notes = df_daily_notes[~((df_daily_notes['階段名稱'] == stage_choice) & (df_daily_notes['日期'].astype(str).isin(edited_dates)))]
        final_notes = pd.concat([other_notes, save_df], ignore_index=True)
        _save_daily_notes_sorted(final_notes)
        st.rerun()

    st.divider()
    st.markdown(f"#### 🗺️ 【{stage_choice}】單階段專用地圖（僅顯示本階段挖掘進度）")
    st.markdown("⬜ 尚未開始 🟧 進行中 🟩 已完成")

    df_logs_for_map = load_sheet_data("dispatch_logs")
    zone_vol_dict = {}
    if not df_logs_for_map.empty and '出土分區' in df_logs_for_map.columns and '載運方量(m³)' in df_logs_for_map.columns:
        tmp_map_logs = df_logs_for_map[df_logs_for_map['出土分區'] != '未指定'].copy()
        tmp_map_logs['載運方量(m³)'] = pd.to_numeric(tmp_map_logs['載運方量(m³)'], errors='coerce')
        zone_vol_dict = tmp_map_logs.groupby('出土分區')['載運方量(m³)'].sum().to_dict()

    stage_idx = get_stage_index(stage_choice)
    stage_thresholds_dict = df_results.set_index('分區代號')['各階累計方量'].to_dict() if not df_results.empty else {}

    fig_stage = go.Figure()
    for idx, row in df_stage_map.iterrows():
        grid_id = row['分區代號']
        current_vol = zone_vol_dict.get(grid_id, 0)
        thresholds = stage_thresholds_dict.get(grid_id, [])
        t_vol = row[target_col] if target_col else 0

        if stage_idx is not None and thresholds and stage_idx < len(thresholds):
            band_start = thresholds[stage_idx - 1] if stage_idx > 0 else 0
            band_end = thresholds[stage_idx]
            band_size = band_end - band_start
            pct = min(max((current_vol - band_start) / band_size * 100, 0), 100) if band_size > 0 else 100.0

            if pct >= 98:
                fill_color = 'rgba(46, 204, 113, 0.8)'
                status_label = "已完成"
            elif pct > 0:
                fill_color = 'rgba(230, 126, 34, 0.7)'
                status_label = f"進行中 {pct:.0f}%"
            else:
                fill_color = 'rgba(240, 240, 240, 0.5)'
                status_label = "尚未開始"
            hover_text = f"{grid_id}<br>本階段目標: {t_vol:,.0f} m³<br>本階段狀態: {status_label}"
        else:
            fill_color = 'rgba(0, 100, 255, 0.1)'
            hover_text = f"{grid_id}<br>階段基準方量: {t_vol:,.0f} m³<br>（開挖前土方無分區進度）"

        fig_stage.add_trace(go.Scatter(
            x=[row['x_min'], row['x_max'], row['x_max'], row['x_min'], row['x_min']],
            y=[row['y_min'], row['y_min'], row['y_max'], row['y_max'], row['y_min']],
            mode='lines', line=dict(color='gray', width=1),
            fill='toself', fillcolor=fill_color, showlegend=False, hoverinfo='text',
            text=hover_text
        ))
        fig_stage.add_annotation(x=row['x_center'], y=row['y_center'], text=grid_id, showarrow=False, font=dict(color="black", size=10))

    fig_stage.update_layout(title=f"【{stage_choice}】單階段進度地圖", dragmode='pan', xaxis_title="", yaxis_title="", yaxis=dict(scaleanchor="x", scaleratio=1), height=500, margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig_stage, use_container_width=True, config={'displayModeBar': False})

with tab_sync:
    st.write("### 🧾 官方聯單時間序列精準對帳與校正")
    st.info("💡 演算法說明：系統會自動尋找時間最接近的紀錄綁定並寫入聯單序號（保留分區），多出的自動剔除，少按的會依官方時序自動補齊。上傳CSV後按一次按鈕即可完成比對並直接寫入雲端，不需要再分兩步驟。")

    sync_date = st.date_input("選擇對帳日期：", value=(datetime.utcnow() + timedelta(hours=8)).date())

    uploaded_csv = st.file_uploader("上傳官方電子聯單 CSV 檔案", type=["csv"])

    if uploaded_csv:
        try:
            try:
                official_df = pd.read_csv(uploaded_csv, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_csv.seek(0)
                official_df = pd.read_csv(uploaded_csv, encoding='big5')

            plate_col = "出場車頭車號"
            datetime_col = "出場日期"
            serial_col = "聯單序號"

            missing_cols = [c for c in [plate_col, datetime_col, serial_col] if c not in official_df.columns]

            if missing_cols:
                st.error(f"⚠️ CSV 檔案格式不符！缺少必要欄位：{', '.join(missing_cols)}。請確認官方匯出的檔案是否包含這些欄位。")
            else:
                if st.button("🚀 開始比對並直接寫入雲端", use_container_width=True):
                    with st.spinner("比對中並同步寫入雲端，請稍候..."):
                        official_df['FullTime'] = pd.to_datetime(official_df[datetime_col], errors='coerce')
                        official_df['ParsedDate'] = official_df['FullTime'].dt.date
                        official_df['正規化車號'] = official_df[plate_col].astype(str).str.replace(r'\W+', '', regex=True).str.upper()

                        sync_off_df = official_df[official_df['ParsedDate'] == sync_date].copy()

                        df_logs = load_sheet_data("dispatch_logs")
                        if df_logs.empty:
                            df_logs = pd.DataFrame(columns=["日期", "時間", "車頭車號", "出土分區", "載運方量(m³)", "備註", "聯單序號"])
                        if "聯單序號" not in df_logs.columns:
                            df_logs["聯單序號"] = ""

                        df_logs['ParsedDate'] = pd.to_datetime(df_logs['日期']).dt.date
                        df_logs['FullTime'] = pd.to_datetime(df_logs['日期'].astype(str) + ' ' + df_logs['時間'].astype(str), errors='coerce')
                        df_logs['正規化車號'] = df_logs['車頭車號'].astype(str).str.replace(r'\W+', '', regex=True).str.upper()
                        sync_sys_df = df_logs[df_logs['ParsedDate'] == sync_date].copy()

                        # 先算出差異摘要，等下比對完一起顯示給你看
                        off_counts = sync_off_df['正規化車號'].value_counts().reset_index()
                        off_counts.columns = ['車頭車號', '官方台數']
                        sys_counts = sync_sys_df['正規化車號'].value_counts().reset_index()
                        sys_counts.columns = ['車頭車號', '系統台數']
                        merged = pd.merge(off_counts, sys_counts, on='車頭車號', how='outer').fillna(0)
                        merged['差異 (多按或漏按)'] = merged['系統台數'] - merged['官方台數']

                        # 直接執行時間序列精準比對覆蓋與序號寫入
                        to_delete_indices = []
                        to_add_records = []
                        updates = {}

                        plates = set(sync_off_df['正規化車號']).union(set(sync_sys_df['正規化車號']))
                        for plate in plates:
                            o_subset = sync_off_df[sync_off_df['正規化車號'] == plate].sort_values('FullTime')
                            s_subset = df_logs[(df_logs['ParsedDate'] == sync_date) & (df_logs['正規化車號'] == plate)].sort_values('FullTime')

                            s_indices = s_subset.index.tolist()
                            used_s = set()

                            for _, o_row in o_subset.iterrows():
                                o_t = o_row['FullTime']
                                o_plate_raw = o_row[plate_col]
                                o_serial_raw = str(o_row[serial_col])

                                best_s_idx = None
                                best_diff = float('inf')

                                for idx in s_indices:
                                    if idx in used_s: continue
                                    st_t = s_subset.loc[idx, 'FullTime']
                                    if pd.isna(o_t) or pd.isna(st_t): continue
                                    diff = abs((o_t - st_t).total_seconds())
                                    if diff < best_diff:
                                        best_diff = diff
                                        best_s_idx = idx

                                if best_s_idx is not None and best_diff < 7200:
                                    used_s.add(best_s_idx)
                                    updates[best_s_idx] = {
                                        "時間": o_t.strftime("%H:%M:%S"),
                                        "聯單序號": o_serial_raw
                                    }
                                else:
                                    to_add_records.append({
                                        "日期": o_t.strftime("%Y-%m-%d") if pd.notnull(o_t) else sync_date.strftime("%Y-%m-%d"),
                                        "時間": o_t.strftime("%H:%M:%S") if pd.notnull(o_t) else "00:00:00",
                                        "車頭車號": o_plate_raw,
                                        "出土分區": "未指定",
                                        "載運方量(m³)": 12.0,
                                        "備註": "官方聯單補登",
                                        "聯單序號": o_serial_raw
                                    })

                            for idx in s_indices:
                                if idx not in used_s:
                                    to_delete_indices.append(idx)

                        if updates:
                            for idx, vals in updates.items():
                                for k, v in vals.items():
                                    df_logs.loc[idx, k] = v

                        if to_delete_indices:
                            df_logs = df_logs.drop(index=to_delete_indices)

                        if to_add_records:
                            df_logs = pd.concat([df_logs, pd.DataFrame(to_add_records)], ignore_index=True)

                        df_logs['SortTime'] = pd.to_datetime(df_logs['日期'].astype(str) + ' ' + df_logs['時間'].astype(str), errors='coerce')
                        df_logs = df_logs.sort_values('SortTime').drop(columns=['ParsedDate', 'FullTime', '正規化車號', 'SortTime'])

                        save_ok = save_sheet_data("dispatch_logs", df_logs)

                    if save_ok:
                        st.success(f"✅ 比對完成並已直接寫入雲端！更新 {len(updates)} 筆、新增 {len(to_add_records)} 筆、剔除 {len(to_delete_indices)} 筆。")
                        st.markdown("##### 比對差異摘要（僅供參考，資料已寫入雲端）")
                        st.dataframe(merged, use_container_width=True)
                    else:
                        st.error("寫入雲端失敗，請檢查連線或稍後再試。")

        except Exception as e:
            st.error(f"檔案解析或比對失敗：{e}")

with tab_manifest:
    st.write("### 🎫 聯單庫存與發放管理")

    df_manifest = load_sheet_data("manifest_settings")
    if df_manifest.empty:
        df_manifest = pd.DataFrame({
            "聯單類型": ["B1", "B2-3", "B4", "B5"],
            "總配額": [1000, 2790, 2821, 30],
            "已列印數量": [0, 0, 0, 0]
        })

    df_logs = load_sheet_data("dispatch_logs")

    used_counts = {t: 0 for t in df_manifest["聯單類型"]}
    if not df_logs.empty and "聯單序號" in df_logs.columns:
        valid_serials = df_logs["聯單序號"].dropna().astype(str).str.strip().str.upper()
        valid_serials = valid_serials[(valid_serials != "") & (valid_serials != "NAN") & (valid_serials != "NONE")]

        for m_type in df_manifest["聯單類型"]:
            count = valid_serials.str.contains(m_type.upper(), regex=False).sum()
            used_counts[m_type] = count

    df_manifest["總配額"] = pd.to_numeric(df_manifest["總配額"], errors='coerce').fillna(0).astype(int)
    col_print_name = "聯單數量_已列印" if "聯單數量_已列印" in df_manifest.columns else "已列印數量"
    df_manifest["已列印數量"] = pd.to_numeric(df_manifest[col_print_name], errors='coerce').fillna(0).astype(int)
    df_manifest["已使用數量"] = df_manifest["聯單類型"].map(used_counts).fillna(0).astype(int)
    df_manifest["現場剩餘可用"] = df_manifest["已列印數量"] - df_manifest["已使用數量"]
    df_manifest["雲端未列印配額"] = df_manifest["總配額"] - df_manifest["已列印數量"]

    st.info("請於下方表格直接修改「已列印數量」，系統會根據對帳結果自動計算現場剩餘的可用張數。")
    edited_manifest = st.data_editor(
        df_manifest,
        column_config={
            "總配額": st.column_config.NumberColumn(format="%d"),
            "已列印數量": st.column_config.NumberColumn(format="%d", min_value=0),
            "已使用數量": st.column_config.NumberColumn(format="%d"),
            "現場剩餘可用": st.column_config.NumberColumn(format="%d"),
            "雲端未列印配額": st.column_config.NumberColumn(format="%d"),
        },
        disabled=["聯單類型", "總配額", "已使用數量", "現場剩餘可用", "雲端未列印配額"],
        hide_index=True,
        use_container_width=True
    )

    if st.button("💾 儲存列印數量更新"):
        if save_sheet_data("manifest_settings", edited_manifest[['聯單類型', '總配額', '已列印數量']]):
            st.success("已更新列印數量！")
            st.rerun()

    st.markdown("#### 🚨 現場庫存狀態警報")

    alert_triggered = False
    for idx, row in df_manifest.iterrows():
        if pd.notnull(row["現場剩餘可用"]) and pd.notnull(row["雲端未列印配額"]):
            if row["現場剩餘可用"] < 100 and row["雲端未列印配額"] > 0:
                st.error(f"⚠️ **【警告】{row['聯單類型']}** 聯單現場僅剩 **{int(row['現場剩餘可用'])}** 張！請盡速列印補充備用。 (尚有雲端配額 {int(row['雲端未列印配額'])} 張)")
                alert_triggered = True

    if not alert_triggered:
        st.success("✅ 目前所有類型的聯單現場庫存皆十分充足，或已全數列印完畢。")

with tab_delivery:
    st.write("### ✍️ 現場廠商交付簽收管理")

    df_delivery = load_sheet_data("manifest_delivery")
    if df_delivery.empty:
        df_delivery = pd.DataFrame(columns=["交付日期", "交付時間", "廠商名稱", "聯單類型", "起始序號", "發放張數", "簽收人姓名", "簽名資料"])

    if not df_delivery.empty:
        try:
            st.markdown("#### 📋 歷史交付簽收對帳看板")

            record_options = [f"[{r['交付日期']} {r['交付時間']}] {r['廠商名稱']} {r['簽收人姓名']} ({r['聯單類型']} 聯單 / {int(r['發放張數']) if pd.notnull(r['發放張數']) else 0}張)" for idx, r in df_delivery.iterrows()]
            selected_record_idx = st.selectbox("🔍 選擇一筆歷史紀錄：", options=range(len(record_options)), format_func=lambda x: record_options[x], key="select_history_delivery")

            chosen_row = df_delivery.iloc[selected_record_idx]

            clean_history_serial = str(chosen_row['起始序號']).strip()
            if clean_history_serial.endswith('.0'):
                clean_history_serial = clean_history_serial[:-2]

            st.info(f"""**詳細交付核對資訊：**
* 交付時間： `{chosen_row['交付日期']} {chosen_row['交付時間']}`
* 簽收廠商： `{chosen_row['廠商名稱']}`
* 聯單規格： `{chosen_row['聯單類型']}`
* 發放張數： `{int(chosen_row['發放張數']) if pd.notnull(chosen_row['發放張數']) else 0} 張`
* 起始序號： `{clean_history_serial}`
* 現場簽收人： `{chosen_row['簽收人姓名']}`

*(簽名影像已隱藏，請匯出 PDF 檢視)*""")

            st.divider()

            st.markdown("#### 📄 導出交付簽收 PDF 報表")
            col_pdf1, col_pdf2 = st.columns([2, 1])
            with col_pdf1:
                pdf_scope = st.selectbox("選擇要匯出的聯單範圍：", options=["全部聯單類型", "B1", "B2-3", "B4", "B5"], key="pdf_scope_select")
            with col_pdf2:
                st.write("")
                st.write("")
                if pdf_scope == "全部聯單類型":
                    df_pdf_target = df_delivery.copy()
                else:
                    df_pdf_target = df_delivery[df_delivery["聯單類型"] == pdf_scope].copy()

                if df_pdf_target.empty:
                    st.warning("⚠️ 該範圍內無任何交付紀錄，無法產生報表。")
                else:
                    if st.button("📥 下載 PDF 報表", key="btn_download_delivery_pdf"):
                        with st.spinner("產生中..."):
                            try:
                                delivery_pdf_path = generate_delivery_pdf(df_pdf_target, pdf_scope)
                                with open(delivery_pdf_path, "rb") as pdf_file:
                                    st.download_button(
                                        label=f"✅ 點此儲存 {pdf_scope} 簽收報表 (含影像)",
                                        data=pdf_file,
                                        file_name=f"delivery_report_{pdf_scope}_{datetime.now().strftime('%Y%m%d')}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                            except Exception as ex:
                                st.error(f"PDF 導出失敗：{ex}")
            st.divider()
        except Exception as e:
            st.error(f"⚠️ 歷史交付看板渲染發生錯誤（已略過，其餘功能不受影響）：{e}")

    st.markdown("#### 📥 新增聯單現場發放")

    input_type = st.selectbox("選擇本次發放的聯單類型", options=["B1", "B2-3", "B4", "B5"], key="delivery_type_select")

    df_manifest_check = load_sheet_data("manifest_settings")
    df_logs_check = load_sheet_data("dispatch_logs")

    available_stock = 0
    if not df_manifest_check.empty and "聯單類型" in df_manifest_check.columns:
        col_target = "聯單數量_已列印" if "聯單數量_已列印" in df_manifest_check.columns else "已列印數量"
        if col_target in df_manifest_check.columns:
            used_counts_check = {t: 0 for t in df_manifest_check["聯單類型"]}
            if not df_logs_check.empty and "聯單序號" in df_logs_check.columns:
                valid_serials_check = df_logs_check["聯單序號"].dropna().astype(str).str.strip().str.upper()
                valid_serials_check = valid_serials_check[(valid_serials_check != "") & (valid_serials_check != "NAN") & (valid_serials_check != "NONE")]
                for m_type in df_manifest_check["聯單類型"]:
                    count = valid_serials_check.str.contains(m_type.upper(), regex=False).sum()
                    used_counts_check[m_type] = count

            df_manifest_check["已使用數量"] = df_manifest_check["聯單類型"].map(used_counts_check).fillna(0).astype(int)
            df_manifest_check["現場剩餘可用"] = pd.to_numeric(df_manifest_check[col_target], errors='coerce').fillna(0).astype(int) - df_manifest_check["已使用數量"]

            match_stock_row = df_manifest_check[df_manifest_check["聯單類型"] == input_type]
            if not match_stock_row.empty and "現場剩餘可用" in match_stock_row.columns:
                available_stock = int(match_stock_row.iloc[0]["現場剩餘可用"])

    st.write(f"📊 該類型聯單目前雲端剩餘可交付數量： **{available_stock}** 張")

    auto_serial = ""
    if not df_delivery.empty and "聯單類型" in df_delivery.columns and "起始序號" in df_delivery.columns and "發放張數" in df_delivery.columns:
        df_type_last = df_delivery[df_delivery["聯單類型"] == input_type]
        if not df_type_last.empty:
            last_record = df_type_last.iloc[-1]
            last_serial = str(last_record["起始序號"]).strip()
            if last_serial.endswith('.0'):
                last_serial = last_serial[:-2]
            try:
                last_count = int(last_record["發放張數"]) if pd.notnull(last_record["發放張數"]) else 0
                match = re.search(r'\d+', last_serial)
                if match:
                    num_str = match.group()
                    prefix = last_serial[:match.start()]
                    suffix = last_serial[match.end():]
                    next_num = int(num_str) + last_count
                    auto_serial = f"{prefix}{next_num:0{len(num_str)}d}{suffix}"
            except Exception:
                auto_serial = ""

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        input_vendor = st.text_input("廠商名稱", value="力勤", disabled=True)
        input_count = st.number_input("發放張數", min_value=1, value=50, step=1, format="%d")
    with col_d2:
        input_serial = st.text_input("聯單起始序號 (可根據現場實物修改)", value=auto_serial)
        input_name = st.text_input("廠商簽收人姓名", value="")

    sign_active = st.checkbox("✍️ 填寫完畢後，點此展開簽名板")

    base64_sign = ""
    if sign_active:
        from streamlit_drawable_canvas import st_canvas
        from PIL import Image
        import io
        import base64

        canvas_sign = st_canvas(
            fill_color="rgba(255, 255, 255, 1)",
            stroke_width=3,
            stroke_color="#000000",
            background_color="#EEEEEE",
            height=200,
            width=340,
            drawing_mode="freedraw",
            update_streamlit=True,
            key=f"canvas_delivery_form_{st.session_state['canvas_key_counter']}",
        )

        if st.button("確認交付並永久儲存簽收紀錄", use_container_width=True):
            if not input_serial.strip():
                st.error("請填寫聯單起始序號")
            elif not input_name.strip():
                st.error("請填寫廠商簽收人姓名")
            elif canvas_sign.image_data is None or np.sum(canvas_sign.image_data[:, :, 3]) == 0:
                st.error("請完成手寫簽名後再提交")
            elif int(input_count) > available_stock:
                st.error(f"❌ 拒絕紀錄：庫存量不足！目前庫存僅剩 {available_stock} 張，無法超量發放 {int(input_count)} 張。")
            else:
                img_array = canvas_sign.image_data.astype('uint8')
                img = Image.fromarray(img_array, 'RGBA')
                buffered = io.BytesIO()
                img.save(buffered, format="PNG")
                base64_sign = base64.b64encode(buffered.getvalue()).decode()

                clean_input_serial = input_serial.strip()
                if clean_input_serial.endswith('.0'):
                    clean_input_serial = clean_input_serial[:-2]

                now_tw = datetime.utcnow() + timedelta(hours=8)
                new_record = {
                    "交付日期": now_tw.strftime("%Y-%m-%d"),
                    "交付時間": now_tw.strftime("%H:%M:%S"),
                    "廠商名稱": input_vendor.strip(),
                    "聯單類型": input_type,
                    "起始序號": clean_input_serial,
                    "發放張數": int(input_count),
                    "簽收人姓名": input_name.strip(),
                    "簽名資料": base64_sign
                }

                df_delivery = pd.concat([df_delivery, pd.DataFrame([new_record])], ignore_index=True)

                if save_sheet_data("manifest_delivery", df_delivery):
                    st.session_state['canvas_key_counter'] += 1
                    st.success("✅ 聯單現場交付成功！紀錄與簽名已即時寫入雲端。")
                    st.rerun()

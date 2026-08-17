import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import re
import os
import io
import json
import tempfile
import datetime
import xlsxwriter.utility
import gspread
from oauth2client.service_account import ServiceAccountCredentials

try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_READY = True
except ImportError:
    MATPLOTLIB_READY = False

try:
    from adjustText import adjust_text
    ADJUSTTEXT_READY = True
except ImportError:
    ADJUSTTEXT_READY = False

try:
    import ezdxf
    EZDXF_READY = True
except ImportError:
    EZDXF_READY = False

st.set_page_config(page_title="樁位進度管理系統 (多工地版)", layout="wide")
st.title("🏗️ 樁位進度管理系統")

# ============================================================
# 基礎 session_state
# ============================================================
for key, default in [
    ('sel_a', []), ('sel_b', []),
    ('site_dxf_cache', {}), ('site_boundary_cache', {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default

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

# ============================================================
# Google Sheets 連線 (只開試算表本身，不綁定固定工作表)
# ============================================================
def _load_gs_secrets():
    """
    支援三種 secrets 寫法，自動判斷使用哪一種：
    1) [connections.gsheets] 攤平格式：
         [connections.gsheets]
         spreadsheet = "https://docs.google.com/..."
         type = "service_account"
         private_key = "..."
         ...(其餘服務帳號欄位)
    2) [gcp_service_account] TOML 表格 (推薦，最不容易踩換行/跳脫字元的雷)：
         sheet_url = "https://docs.google.com/..."
         [gcp_service_account]
         type = "service_account"
         private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
         ...
    3) 舊版：sheet_url + gcp_service_account = 整包 JSON 字串包在三引號內 (最容易因換行格式出錯，不建議)
    回傳 (creds_dict, sheet_url)
    """
    try:
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            cfg = dict(st.secrets["connections"]["gsheets"])
            sheet_url = cfg.pop("spreadsheet", None) or cfg.pop("sheet_url", None)
            if sheet_url:
                return cfg, sheet_url
    except Exception:
        pass
    try:
        if "gcp_service_account" in st.secrets and not isinstance(st.secrets["gcp_service_account"], str):
            creds_dict = dict(st.secrets["gcp_service_account"])
            sheet_url = st.secrets["sheet_url"]
            if creds_dict and sheet_url:
                return creds_dict, sheet_url
    except Exception:
        pass
    creds_dict = json.loads(st.secrets["gcp_service_account"])
    sheet_url = st.secrets["sheet_url"]
    return creds_dict, sheet_url

@st.cache_resource
def get_gs_client():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict, sheet_url = _load_gs_secrets()
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        ss = client.open_by_url(sheet_url)
        return ss
    except Exception as e:
        st.error(f"雲端連線異常: {e}")
        return None

ss = get_gs_client()

def get_ws(ss, name, header):
    """取得指定名稱的工作表，不存在就自動建立並寫入表頭"""
    if ss is None:
        return None
    try:
        return ss.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=500, cols=max(10, len(header)))
        ws.append_row(header)
        return ws
    except Exception as e:
        st.error(f"工作表存取失敗 ({name}): {e}")
        return None

# ============================================================
# 工地清單 (每個工地一列設定資料)
# ============================================================
SITE_LIST_WS = "工地清單"
SITE_COLUMNS = [
    'site_id', 'site_name', 'total_piles',
    'dxf_pile_source', 'dxf_pile_block', 'dxf_boundary_layer', 'match_radius',
    'pdf_loc_note_right', 'pdf_loc_note_left', 'pdf_week_est',
    'dxf_drive_file_id'
]
SITE_DEFAULTS = {
    'total_piles': 262, 'dxf_pile_source': 'block', 'dxf_pile_block': 'HH3', 'dxf_boundary_layer': '開挖邊界',
    'match_radius': 800, 'pdf_loc_note_right': '', 'pdf_loc_note_left': '', 'pdf_week_est': 20,
    'dxf_drive_file_id': ''
}
PILE_SOURCE_LABELS = {
    'block': '圖塊 (INSERT Block，未分解)',
    'circle': '圓形 (CIRCLE，分解後常見)',
    'point': '點 (POINT，分解後常見)',
    'text': '文字本身位置 (沒有另外的幾何符號，直接拿文字座標當樁位)'
}

@st.cache_data(ttl=60)
def load_sites(_ss):
    ws = get_ws(_ss, SITE_LIST_WS, SITE_COLUMNS)
    if ws is None:
        return pd.DataFrame(columns=SITE_COLUMNS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=SITE_COLUMNS)
    df = pd.DataFrame(records)
    for c in SITE_COLUMNS:
        if c not in df.columns:
            df[c] = SITE_DEFAULTS.get(c, '')
    return df

def add_site(ss, site_dict):
    ws = get_ws(ss, SITE_LIST_WS, SITE_COLUMNS)
    if ws is None:
        return
    header = ws.row_values(1)
    changed = False
    for c in SITE_COLUMNS:
        if c not in header:
            header.append(c)
            changed = True
    if changed:
        ws.update('A1', [header])
    row = [str(site_dict.get(c, SITE_DEFAULTS.get(c, ''))) for c in header]
    ws.append_row(row)
    st.cache_data.clear()

def update_site_field(ss, site_id, field, value):
    """更新工地清單裡指定工地的某一欄位 (依實際表頭欄位定位，避免新舊表格欄位順序不一致而寫錯欄)"""
    ws = get_ws(ss, SITE_LIST_WS, SITE_COLUMNS)
    if ws is None:
        return
    header = ws.row_values(1)
    if field not in header:
        header.append(field)
        ws.update('A1', [header])
    col_num = header.index(field) + 1
    records = ws.get_all_records()
    for i, r in enumerate(records):
        if str(r.get('site_id')) == str(site_id):
            row_num = i + 2  # +1 header, +1 一位起算
            ws.update_cell(row_num, col_num, str(value))
            break
    st.cache_data.clear()

df_sites = load_sites(ss)

st.markdown("### 🏢 選擇工地")
site_names = df_sites['site_name'].tolist() if not df_sites.empty else []
choice_options = site_names + (["➕ 新增工地"] if not demo_mode else [])

if not choice_options:
    st.warning("目前雲端還沒有任何工地資料。")
    if demo_mode:
        st.info("請輸入管理員密碼後即可新增第一個工地。")
        st.stop()
    else:
        choice_options = ["➕ 新增工地"]

sel_site_name = st.selectbox("工地", choice_options, key="site_selector")

if sel_site_name == "➕ 新增工地":
    st.markdown("#### ➕ 新增工地設定")
    with st.form("new_site"):
        n_id = st.text_input("工地代碼 (英數字，如 CDC / TPE01，用於雲端分頁命名，建立後不可更改)")
        n_name = st.text_input("工地顯示名稱 (如：CDC中間樁與共構樁)")
        n_total = st.number_input("樁位總支數", 1, 5000, 262)
        n_source = st.selectbox("樁位在DXF裡是用什麼畫的？", list(PILE_SOURCE_LABELS.keys()), format_func=lambda k: PILE_SOURCE_LABELS[k])
        n_block = st.text_input("圖塊名稱 / 圖層名稱 (依上面選的類型而定)", value="HH3",
                                 help="選「圖塊」：填CAD裡樁位符號的Block名稱。\n選「圓形」或「點」(通常是圖塊分解後的情況)：填這些圓形/點所在的圖層(Layer)名稱。\n選「文字本身位置」：這欄不用填。若用舊版CSV匯入，對應原本 名稱=='HH3' 那個篩選值。")
        n_layer = st.text_input("DXF 開挖邊界圖層(Layer)名稱 (沒有邊界線可留空)", value="開挖邊界",
                                 help="CAD 裡畫開挖邊界線的圖層名稱，程式會抓這個圖層裡的封閉多邊形。留空就不畫邊界。")
        n_radius = st.number_input("樁號文字比對半徑 (CAD單位)", 10, 100000, 800,
                                    help="樁位符號會抓「距離內最近的文字」當作樁號，這裡設定搜尋半徑。")
        n_right = st.text_input("PDF 右側標題預設文字", value="")
        n_left = st.text_input("PDF 左側標題預設文字", value="")
        n_week = st.number_input("本週預計完成支數 (預設值)", 0, 1000, 20)
        if st.form_submit_button("建立工地"):
            if not n_id or not n_name:
                st.error("工地代碼與名稱不可空白")
            elif not df_sites.empty and n_id in df_sites['site_id'].astype(str).values:
                st.error("這個工地代碼已經存在了")
            else:
                add_site(ss, {
                    'site_id': n_id, 'site_name': n_name, 'total_piles': int(n_total),
                    'dxf_pile_source': n_source, 'dxf_pile_block': n_block, 'dxf_boundary_layer': n_layer, 'match_radius': int(n_radius),
                    'pdf_loc_note_right': n_right, 'pdf_loc_note_left': n_left, 'pdf_week_est': int(n_week)
                })
                st.success(f"✅ 工地「{n_name}」已建立，重新整理頁面後即可從下拉選單選取。")
                st.rerun()
    st.stop()

site_row = df_sites[df_sites['site_name'] == sel_site_name].iloc[0]
site_id = str(site_row['site_id'])
TOTAL_PILES = int(site_row['total_piles']) if str(site_row['total_piles']).strip() else 262
DXF_PILE_SOURCE = str(site_row.get('dxf_pile_source', '') or 'block').strip() or 'block'
DXF_PILE_BLOCK = str(site_row['dxf_pile_block']) if str(site_row.get('dxf_pile_block', '')).strip() else 'HH3'
DXF_BOUNDARY_LAYER = str(site_row.get('dxf_boundary_layer', '') or '').strip()
MATCH_RADIUS = float(site_row['match_radius']) if str(site_row['match_radius']).strip() else 800

st.caption(f"📍 目前工地：**{sel_site_name}** (代碼: {site_id})　樁位總支數：{TOTAL_PILES}")

if not demo_mode:
    with st.expander("⚙️ 編輯此工地的 DXF 讀取設定"):
        with st.form(f"edit_site_{site_id}"):
            e_total = st.number_input("樁位總支數", 1, 5000, TOTAL_PILES)
            e_source = st.selectbox("樁位在DXF裡是用什麼畫的？", list(PILE_SOURCE_LABELS.keys()),
                                     index=list(PILE_SOURCE_LABELS.keys()).index(DXF_PILE_SOURCE) if DXF_PILE_SOURCE in PILE_SOURCE_LABELS else 0,
                                     format_func=lambda k: PILE_SOURCE_LABELS[k])
            e_block = st.text_input("圖塊名稱 / 圖層名稱 (依上面選的類型而定，選「文字本身位置」則不用填)", value=DXF_PILE_BLOCK)
            e_layer = st.text_input("DXF 開挖邊界圖層(Layer)名稱 (沒有邊界線可留空)", value=DXF_BOUNDARY_LAYER)
            e_radius = st.number_input("樁號文字比對半徑 (CAD單位)", 10, 100000, int(MATCH_RADIUS))
            if st.form_submit_button("💾 儲存設定"):
                update_site_field(ss, site_id, 'total_piles', e_total)
                update_site_field(ss, site_id, 'dxf_pile_source', e_source)
                update_site_field(ss, site_id, 'dxf_pile_block', e_block)
                update_site_field(ss, site_id, 'dxf_boundary_layer', e_layer)
                update_site_field(ss, site_id, 'match_radius', e_radius)
                st.success("✅ 已更新，重新整理套用新設定")
                st.rerun()

HIST_WS_NAME = f"{site_id}_施工明細"
SETTINGS_WS_NAME = f"{site_id}_系統設定"
EXTRA_WS_NAME = f"{site_id}_額外樁位"
HIST_COLUMNS = ['樁號', '施工日期', '機台', '施作順序', 'X', 'Y']
EXTRA_COLUMNS = ['樁號', 'X', 'Y']

# ============================================================
# 樁位/邊界資料來源：DXF 自動讀取 或 舊版 CSV 上傳
# ============================================================
def _load_dxf_doc(file_bytes):
    """
    容錯讀取 DXF：優先用標準讀取，失敗或發現結構問題就改用 ezdxf.recover
    (Revit/Civil3D 等第三方軟體匯出的DXF常有非標準結構，recover模式會自動修復再讀取)
    """
    with tempfile.NamedTemporaryFile(suffix='.dxf', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        try:
            doc = ezdxf.readfile(tmp_path)
        except Exception:
            from ezdxf import recover
            doc, auditor = recover.readfile(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    return doc

@st.cache_data(show_spinner="🔍 掃描 DXF 內容中...")
def scan_dxf_summary(file_bytes):
    """掃描DXF，列出每個圖層上各種圖元的數量、以及每個圖層上INSERT圖塊的名稱與數量，供使用者對照設定"""
    doc = _load_dxf_doc(file_bytes)
    msp = doc.modelspace()
    layer_entity_counter = {}
    insert_block_counter = {}
    for e in msp:
        try:
            layer = e.dxf.layer
        except Exception:
            layer = '(無圖層)'
        et = e.dxftype()
        key = (layer, et)
        layer_entity_counter[key] = layer_entity_counter.get(key, 0) + 1
        if et == 'INSERT':
            try:
                bkey = (layer, e.dxf.name)
                insert_block_counter[bkey] = insert_block_counter.get(bkey, 0) + 1
            except Exception:
                pass
    df_layer = pd.DataFrame(
        [{'圖層': k[0], '圖元類型': k[1], '數量': v} for k, v in layer_entity_counter.items()]
    ).sort_values(['圖層', '數量'], ascending=[True, False])
    df_block = pd.DataFrame(
        [{'圖層': k[0], '圖塊名稱': k[1], '數量': v} for k, v in insert_block_counter.items()]
    ).sort_values('數量', ascending=False)
    return df_layer, df_block

def parse_dxf(file_bytes, pile_source, pile_block, boundary_layer, match_radius=800):
    """
    讀取 DXF：
    - pile_source == 'block' : 抓指定「圖塊(Block)」的插入點當樁位 (pile_block = 圖塊名稱，支援部分包含比對)
    - pile_source == 'circle': 抓指定「圖層」上的圓形(CIRCLE)圓心當樁位 (pile_block = 圖層名稱；留空則不篩選圖層)
    - pile_source == 'point' : 抓指定「圖層」上的點(POINT)當樁位 (pile_block = 圖層名稱；留空則不篩選圖層)
    - pile_source == 'text'  : 沒有另外的幾何符號，直接把文字本身的座標當樁位
    再抓「距離最近的文字」當樁號 (text 模式除外，樁號就是文字內容本身)。
    boundary_layer 留空則不畫開挖邊界。
    """
    doc = _load_dxf_doc(file_bytes)
    msp = doc.modelspace()

    texts = []
    for e in msp.query('TEXT MTEXT'):
        try:
            if e.dxftype() == 'MTEXT':
                content = e.plain_text() if hasattr(e, 'plain_text') else e.text
                ins = e.dxf.insert
            else:
                content = e.dxf.text
                ins = e.dxf.insert
            texts.append((ins.x, ins.y, str(content)))
        except Exception:
            continue
    texts_df = pd.DataFrame(texts, columns=['X', 'Y', '內容'])

    pile_block = (pile_block or '').strip()

    if pile_source == 'text':
        # 沒有幾何符號，文字本身的位置就是樁位，文字內容就是樁號
        if texts_df.empty:
            piles_df = pd.DataFrame(columns=['X', 'Y', '樁號', '數字'])
        else:
            piles_df = texts_df.rename(columns={'內容': '樁號'}).copy()
            piles_df['樁號'] = piles_df['樁號'].astype(str).str.strip().str.upper()
            piles_df = piles_df[piles_df['樁號'].str.contains(r'\d', regex=True, na=False)]
            piles_df['數字'] = piles_df['樁號'].str.extract(r'(\d+)').fillna(0).astype(int)
            piles_df = piles_df.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字').reset_index(drop=True)
    else:
        piles = []
        if pile_source == 'block':
            all_inserts = list(msp.query('INSERT'))
            pb_norm = pile_block.upper()
            exact = [e for e in all_inserts if getattr(e.dxf, 'name', '').strip().upper() == pb_norm]
            candidates = exact if exact else [e for e in all_inserts if pb_norm in getattr(e.dxf, 'name', '').strip().upper()]
            for e in candidates:
                try:
                    ins = e.dxf.insert
                    piles.append((ins.x, ins.y))
                except Exception:
                    continue
        elif pile_source == 'circle':
            for e in msp.query('CIRCLE'):
                try:
                    if (not pile_block) or e.dxf.layer.strip() == pile_block:
                        c = e.dxf.center
                        piles.append((c.x, c.y))
                except Exception:
                    continue
        elif pile_source == 'point':
            for e in msp.query('POINT'):
                try:
                    if (not pile_block) or e.dxf.layer.strip() == pile_block:
                        loc = e.dxf.location
                        piles.append((loc.x, loc.y))
                except Exception:
                    continue

        piles_df = pd.DataFrame(piles, columns=['X', 'Y'])

        def get_nearest_text(px, py):
            if texts_df.empty:
                return "未命名"
            dist = np.sqrt((texts_df['X'] - px) ** 2 + (texts_df['Y'] - py) ** 2)
            idx = dist.idxmin()
            if dist[idx] < match_radius:
                return str(texts_df.loc[idx, '內容']).strip()
            return "未命名"

        if not piles_df.empty:
            piles_df['樁號'] = piles_df.apply(lambda r: get_nearest_text(r['X'], r['Y']), axis=1)
            piles_df['樁號'] = piles_df['樁號'].astype(str).str.strip().str.upper()
            piles_df['數字'] = piles_df['樁號'].str.extract(r'(\d+)').fillna(0).astype(int)
            piles_df = piles_df.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字').reset_index(drop=True)
        else:
            piles_df = pd.DataFrame(columns=['X', 'Y', '樁號', '數字'])

    loops = []
    boundary_layer = (boundary_layer or '').strip()
    if boundary_layer:
        for e in msp.query('LWPOLYLINE'):
            try:
                if e.dxf.layer == boundary_layer:
                    pts = [(p[0], p[1]) for p in e.get_points()]
                    if pts:
                        loops.append(pts)
            except Exception:
                continue
        for e in msp.query('POLYLINE'):
            try:
                if e.dxf.layer == boundary_layer:
                    pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
                    if pts:
                        loops.append(pts)
            except Exception:
                continue

    return piles_df, loops

def parse_legacy_csv(boundary_file, pile_file, pile_block):
    """舊版 CSV 匯入路徑 (相容舊工地資料)"""
    try:
        try:
            df_b = pd.read_csv(boundary_file, encoding='utf-8')
        except Exception:
            boundary_file.seek(0)
            df_b = pd.read_csv(boundary_file, encoding='big5')
        x_col = next((c for c in df_b.columns if 'X' in c.upper() or '座標' in c), None)
        y_col = next((c for c in df_b.columns if 'Y' in c.upper() or '座標' in c), None)
        text_col = next((c for c in df_b.columns if '內容' in c or '值' in c or '樁號' in c), None)
        df_b['樁號'] = df_b[text_col].apply(lambda x: re.sub(r'\\[^;]+;|[{}]', '', str(x)).strip().upper())
        df_b = df_b[df_b['樁號'].str.match(r'^P\d+$')]
        df_b['數字'] = df_b['樁號'].str.extract(r'(\d+)').astype(int)
        df_b['X'] = pd.to_numeric(df_b[x_col], errors='coerce')
        df_b['Y'] = pd.to_numeric(df_b[y_col], errors='coerce')
        df_b = df_b.dropna(subset=['X', 'Y']).sort_values('數字')
        loop_pts = list(zip(df_b['X'], df_b['Y']))
        loops = [loop_pts] if loop_pts else []
    except Exception as e:
        st.error(f"邊界CSV讀取失敗: {e}")
        loops = []

    try:
        try:
            df_p = pd.read_csv(pile_file, encoding='utf-8')
        except Exception:
            pile_file.seek(0)
            df_p = pd.read_csv(pile_file, encoding='big5')
        texts = df_p[df_p['名稱'].isin(['文字', '多行文字']) & df_p['內容'].notnull()][['內容', '位置 X', '位置 Y']].copy()
        texts.rename(columns={'位置 X': 'X', '位置 Y': 'Y'}, inplace=True)
        texts.dropna(subset=['X', 'Y'], inplace=True)

        piles = df_p[df_p['名稱'] == pile_block][['位置 X', '位置 Y']].copy()
        piles.rename(columns={'位置 X': 'X', '位置 Y': 'Y'}, inplace=True)
        piles.dropna(subset=['X', 'Y'], inplace=True)

        def get_nearest_text(px, py):
            if len(texts) == 0:
                return "未命名"
            dist = np.sqrt((texts['X'] - px) ** 2 + (texts['Y'] - py) ** 2)
            idx = dist.idxmin()
            if dist[idx] < 800:
                return str(texts.loc[idx, '內容']).strip()
            return "未命名"

        piles['樁號'] = piles.apply(lambda r: get_nearest_text(r['X'], r['Y']), axis=1)
        piles['樁號'] = piles['樁號'].astype(str).str.strip().str.upper()
        piles['數字'] = piles['樁號'].str.extract(r'(\d+)').fillna(0).astype(int)
        piles_df = piles.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字').reset_index(drop=True)
    except Exception as e:
        st.error(f"樁位CSV讀取失敗: {e}")
        piles_df = pd.DataFrame(columns=['X', 'Y', '樁號', '數字'])

    return piles_df, loops

df_base_raw = st.session_state.site_dxf_cache.get(site_id, pd.DataFrame(columns=['X', 'Y', '樁號', '數字']))
boundary_loops = st.session_state.site_boundary_cache.get(site_id, [])

with st.expander("📐 樁位圖 / 邊界圖 資料來源設定", expanded=df_base_raw.empty):
    st.caption("⚠️ 目前為暫時版本：DXF/CSV 僅存在本次瀏覽器工作階段，重新整理頁面需要重新上傳一次。永久記住(存到GitHub)之後補上。")
    import_mode = st.radio("匯入方式", ["📄 上傳 DXF (自動讀取)", "📊 上傳舊版 CSV (排樁座標.csv + 中間樁.csv)"], horizontal=True, key=f"import_mode_{site_id}")

    if import_mode.startswith("📄"):
        if not EZDXF_READY:
            st.error("尚未安裝 ezdxf 套件，請在 requirements.txt 加入 `ezdxf` 後重新部署。")
        dxf_file = st.file_uploader("上傳工地 DXF 圖檔", type=['dxf'], key=f"dxf_up_{site_id}")
        if dxf_file is not None and EZDXF_READY:
            c_scan, c_parse = st.columns(2)
            if c_scan.button("🔍 掃描這份DXF的圖層/圖塊清單"):
                with st.spinner("掃描中..."):
                    try:
                        df_layer_scan, df_block_scan = scan_dxf_summary(dxf_file.getvalue())
                        st.session_state[f'scan_layer_{site_id}'] = df_layer_scan
                        st.session_state[f'scan_block_{site_id}'] = df_block_scan
                    except Exception as e:
                        st.error(f"掃描失敗: {e}")

            df_layer_scan = st.session_state.get(f'scan_layer_{site_id}')
            df_block_scan = st.session_state.get(f'scan_block_{site_id}')

            if df_layer_scan is not None:
                st.markdown("**各圖層的圖元數量：**")
                st.dataframe(df_layer_scan, use_container_width=True, height=200)

                st.markdown("**直接套用，不用打字/複製貼上 (避免打錯字)：**")

                layer_options = sorted(df_layer_scan['圖層'].unique().tolist())
                cL1, cL2 = st.columns(2)
                with cL1:
                    pick_layer_for_pile = st.selectbox("選一個圖層 → 套用為「圖塊名稱/圖層名稱」欄位 (適用圓形/點模式)", [''] + layer_options, key=f"pick_layer_pile_{site_id}")
                    if pick_layer_for_pile and st.button("✅ 套用為樁位圖層", key=f"apply_layer_pile_{site_id}"):
                        update_site_field(ss, site_id, 'dxf_pile_block', pick_layer_for_pile)
                        st.success(f"已套用：{pick_layer_for_pile}")
                        st.rerun()
                with cL2:
                    pick_layer_for_boundary = st.selectbox("選一個圖層 → 套用為「開挖邊界圖層」", [''] + layer_options, key=f"pick_layer_boundary_{site_id}")
                    if pick_layer_for_boundary and st.button("✅ 套用為邊界圖層", key=f"apply_layer_boundary_{site_id}"):
                        update_site_field(ss, site_id, 'dxf_boundary_layer', pick_layer_for_boundary)
                        st.success(f"已套用：{pick_layer_for_boundary}")
                        st.rerun()

                st.markdown("**各圖層裡的圖塊(INSERT)名稱與數量：**")
                if df_block_scan is None or df_block_scan.empty:
                    st.caption("這份DXF裡沒有任何圖塊(INSERT)，如果樁位是圓形/點，請用上面的圖層清單選取。")
                else:
                    st.dataframe(df_block_scan, use_container_width=True, height=200)
                    block_options = df_block_scan['圖塊名稱'].unique().tolist()
                    pick_block = st.selectbox("選一個圖塊名稱 → 套用為「圖塊名稱」欄位 (適用圖塊模式)", [''] + block_options, key=f"pick_block_{site_id}")
                    if pick_block and st.button("✅ 套用為樁位圖塊", key=f"apply_block_{site_id}"):
                        update_site_field(ss, site_id, 'dxf_pile_block', pick_block)
                        st.success(f"已套用：{pick_block}")
                        st.rerun()

            if c_parse.button("🔄 從 DXF 重新解析樁位與邊界"):
                with st.spinner("解析 DXF 中..."):
                    try:
                        piles_df, loops = parse_dxf(dxf_file.getvalue(), DXF_PILE_SOURCE, DXF_PILE_BLOCK, DXF_BOUNDARY_LAYER, MATCH_RADIUS)
                        st.session_state.site_dxf_cache[site_id] = piles_df
                        st.session_state.site_boundary_cache[site_id] = loops
                        if piles_df.empty:
                            st.warning("⚠️ 解析完成，但讀到 0 支樁位，請先用左邊「掃描」按鈕確認圖層/圖塊名稱有沒有打對。")
                        else:
                            st.success(f"✅ 解析完成：讀到 {len(piles_df)} 支樁位、{len(loops)} 條邊界線")
                        st.rerun()
                    except Exception as e:
                        st.error(f"DXF 解析失敗: {e}")
    else:
        c_up1, c_up2 = st.columns(2)
        boundary_csv = c_up1.file_uploader("排樁座標.csv (開挖邊界)", type=['csv'], key=f"bcsv_{site_id}")
        pile_csv = c_up2.file_uploader("中間樁.csv (樁位)", type=['csv'], key=f"pcsv_{site_id}")
        if boundary_csv is not None and pile_csv is not None:
            if st.button("🔄 從 CSV 重新解析樁位與邊界"):
                with st.spinner("解析 CSV 中..."):
                    piles_df, loops = parse_legacy_csv(boundary_csv, pile_csv, DXF_PILE_BLOCK)
                st.session_state.site_dxf_cache[site_id] = piles_df
                st.session_state.site_boundary_cache[site_id] = loops
                st.success(f"✅ 解析完成：讀到 {len(piles_df)} 支樁位、{len(loops)} 條邊界線")
                st.rerun()

    if not df_base_raw.empty or boundary_loops:
        st.caption(f"目前已載入 {len(df_base_raw)} 支樁位、{len(boundary_loops)} 條邊界線")

# ============================================================
# 額外樁位 (人工補充，取代舊版寫死的座標公式)
# ============================================================
@st.cache_data(ttl=30)
def load_extra_piles(_ss, ws_name):
    ws = get_ws(_ss, ws_name, EXTRA_COLUMNS)
    if ws is None:
        return pd.DataFrame(columns=EXTRA_COLUMNS)
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=EXTRA_COLUMNS)
    return pd.DataFrame(records)

def save_extra_piles(ss, ws_name, df):
    ws = get_ws(ss, ws_name, EXTRA_COLUMNS)
    if ws is None:
        return
    ws.clear()
    rows = [EXTRA_COLUMNS] + df[EXTRA_COLUMNS].astype(str).values.tolist()
    ws.append_rows(rows)
    st.cache_data.clear()

df_extra = load_extra_piles(ss, EXTRA_WS_NAME)

with st.expander("➕ 額外樁位管理 (DXF/CSV 抓不到、需人工補充的樁位，如場外預留樁)", expanded=False):
    edited_extra = st.data_editor(
        df_extra if not df_extra.empty else pd.DataFrame(columns=EXTRA_COLUMNS),
        num_rows="dynamic", use_container_width=True, key=f"extra_editor_{site_id}",
        disabled=demo_mode
    )
    if not demo_mode:
        if st.button("💾 儲存額外樁位設定"):
            save_extra_piles(ss, EXTRA_WS_NAME, edited_extra.dropna(subset=['樁號']))
            st.success("✅ 已儲存")
            st.rerun()
    else:
        st.caption("👀 訪客模式無法編輯，僅供預覽")

if not df_extra.empty:
    df_extra_clean = df_extra.copy()
    df_extra_clean['X'] = pd.to_numeric(df_extra_clean['X'], errors='coerce')
    df_extra_clean['Y'] = pd.to_numeric(df_extra_clean['Y'], errors='coerce')
    df_extra_clean['樁號'] = df_extra_clean['樁號'].astype(str).str.strip().str.upper()
    df_extra_clean = df_extra_clean.dropna(subset=['X', 'Y'])
    df_extra_clean['數字'] = df_extra_clean['樁號'].str.extract(r'(\d+)').fillna(0).astype(int)
    df_base = pd.concat([df_base_raw, df_extra_clean[['X', 'Y', '樁號', '數字']]], ignore_index=True)
else:
    df_base = df_base_raw.copy()

if not df_base.empty:
    df_base = df_base.drop_duplicates(subset=['樁號']).dropna(subset=['X', 'Y']).sort_values('數字').reset_index(drop=True)

# 因為樁位與邊界現在來自同一份 DXF (同一個座標系統)，不再需要舊版那段
# 手動校正 P1 vs 中間樁1號 座標偏移的 hack。若你的樁位/邊界仍分屬不同來源檔案
# 導致座標系統不一致，可在下面自行加回一段座標平移邏輯。

# ============================================================
# 每工地系統設定 (PDF排版/字體大小等)
# ============================================================
def load_settings(ss, ws_name, site_row):
    default_settings = {
        "pdf_loc_note_right": str(site_row.get('pdf_loc_note_right', '')),
        "pdf_loc_note_left": str(site_row.get('pdf_loc_note_left', '')),
        "pdf_week_est": int(site_row['pdf_week_est']) if str(site_row.get('pdf_week_est', '')).strip() else 20,
        "fig_scale": 1.5, "marker_size": 180, "lbl_fontsize": 18, "text_offset": 20,
        "pos_title_y": 0.90, "pos_info_x": 0.05, "pos_info_y": 0.85,
        "pos_loc_x": 0.70, "pos_loc_y": 0.95, "pos_loc_x_left": 0.22, "pos_loc_y_left": 0.55,
        "pos_leg_x": 0.00, "pos_leg_y": 0.00,
        "pos_img_a_x": 0.35, "pos_img_a_y": 0.10, "pos_img_a_w": 0.30,
        "pos_img_b_x": 0.68, "pos_img_b_y": 0.10, "pos_img_b_w": 0.30
    }
    ws = get_ws(ss, ws_name, ['Key', 'Value'])
    if ws is None:
        return default_settings
    try:
        records = ws.get_all_records()
        loaded = {}
        for r in records:
            k = r.get('Key'); v = r.get('Value')
            if k in default_settings:
                if isinstance(default_settings[k], int):
                    loaded[k] = int(float(v))
                elif isinstance(default_settings[k], float):
                    loaded[k] = float(v)
                else:
                    loaded[k] = str(v)
        return {**default_settings, **loaded}
    except Exception:
        return default_settings

def save_settings(ss, ws_name, settings_dict):
    ws = get_ws(ss, ws_name, ['Key', 'Value'])
    if ws is None:
        return
    ws.clear()
    out = [['Key', 'Value']]
    for k, v in settings_dict.items():
        out.append([k, str(v)])
    ws.append_rows(out)

if 'ui_settings' not in st.session_state or st.session_state.get('ui_settings_site') != site_id:
    st.session_state.ui_settings = load_settings(ss, SETTINGS_WS_NAME, site_row)
    st.session_state.ui_settings_site = site_id

s = st.session_state.ui_settings

if "pdf_loc_note_right" not in st.session_state or st.session_state.get('_last_site') != site_id:
    st.session_state["pdf_loc_note_right"] = s['pdf_loc_note_right']
    st.session_state["pdf_loc_note_left"] = s['pdf_loc_note_left']
    st.session_state["pdf_week_est"] = int(s.get('pdf_week_est', 20))
    st.session_state['_last_site'] = site_id

# ============================================================
# 施工明細 (每工地各自一張分頁)
# ============================================================
def fetch_current_data(sh_main):
    if sh_main is None:
        return pd.DataFrame(columns=HIST_COLUMNS)
    try:
        records = sh_main.get_all_records()
        if not records:
            return pd.DataFrame(columns=HIST_COLUMNS)
        df = pd.DataFrame(records)
        df['樁號'] = df['樁號'].astype(str).str.upper().str.strip()
        return df
    except Exception:
        return pd.DataFrame(columns=HIST_COLUMNS)

sh_main = get_ws(ss, HIST_WS_NAME, HIST_COLUMNS)
df_history_cloud = fetch_current_data(sh_main)

if 'df_history_local' not in st.session_state or st.session_state.get('_hist_site') != site_id or not demo_mode:
    st.session_state.df_history_local = df_history_cloud.copy()
    st.session_state['_hist_site'] = site_id

df_history_full = st.session_state.df_history_local if demo_mode else df_history_cloud
if not df_history_full.empty:
    df_history_full['施工日期_DT'] = pd.to_datetime(df_history_full['施工日期'], errors='coerce')

# ============================================================
# 樁號區間解析 (通用版：任何「英文字首+數字」區間都支援)
# ============================================================
def parse_range_generic(raw_str, total_piles):
    plist = []
    if raw_str:
        pts = re.split(r'[,\s]+', raw_str.strip())
        for pt in pts:
            if not pt:
                continue
            if '-' in pt:
                m = re.match(r'^([A-Za-z]*)(\d+)-([A-Za-z]*)(\d+)$', pt.upper())
                if m:
                    p1, n1, p2, n2 = m.groups()
                    prefix = p1 or p2
                    n1, n2 = int(n1), int(n2)
                    if prefix:
                        for n in range(min(n1, n2), max(n1, n2) + 1):
                            plist.append(f"{prefix}{n}")
                    else:
                        if n1 <= n2:
                            for n in range(n1, n2 + 1):
                                plist.append(f"{n}")
                        else:
                            for n in range(n1, total_piles + 1):
                                plist.append(f"{n}")
                            for n in range(1, n2 + 1):
                                plist.append(f"{n}")
            else:
                plist.append(pt.upper())
    return list(dict.fromkeys(plist))

st.markdown("### 📝 進度登錄")
c1, c2, c3 = st.columns([1, 1, 2])
work_date = c1.date_input("日期")
machine = c2.radio("機台", ["A車", "B車"], horizontal=True)
mode = c3.radio("模式", ["4支一循環", "2支一循環", "單支施作"], horizontal=True)
step = 4 if "4支" in mode else (2 if "2支" in mode else 1)

def save_data(piles):
    if not piles:
        return
    m_data = df_history_full[df_history_full['機台'] == machine]
    seq = 0 if m_data.empty else pd.to_numeric(m_data['施作順序'], errors='coerce').max()
    new_d = []
    for p in piles:
        p = p.upper().strip()
        if p not in df_history_full['樁號'].values:
            seq += 1
            if df_base is not None and not df_base.empty:
                b = df_base[df_base['樁號'] == p]
                x, y = (b['X'].iloc[0], b['Y'].iloc[0]) if not b.empty else (0, 0)
            else:
                x, y = 0, 0
            new_d.append([p, str(work_date), machine, int(seq), float(x), float(y)])
    if new_d:
        if demo_mode:
            new_df = pd.DataFrame(new_d, columns=HIST_COLUMNS)
            st.session_state.df_history_local = pd.concat([st.session_state.df_history_local, new_df], ignore_index=True)
            st.toast("👀 沙盒試用：已模擬寫入網頁暫存變數，後台無變動。")
            st.rerun()
        else:
            if sh_main is not None:
                sh_main.append_rows(new_d)
                st.rerun()

def process_and_save(plist):
    if not plist:
        return
    clean_plist = list(dict.fromkeys([p.upper().strip() for p in plist]))
    existing_piles = set(df_history_full['樁號'].values)
    duplicates = [p for p in clean_plist if p in existing_piles]
    if duplicates:
        dup_str = ", ".join(duplicates)
        st.error(f"🛑 **登錄暫停！** 檢測到以下樁號已存在於資料庫中：【 **{dup_str}** 】\n\n為避免資料異常，已暫停本次寫入，請修改確認後再重新登錄。")
    else:
        save_data(clean_plist)

t1, t2 = st.tabs(["🎯 推算", "✏️ 手動"])
with t1:
    with st.form("a"):
        cc1, cc2, cc3 = st.columns(3)
        sp = cc1.number_input("起始號碼", 1, TOTAL_PILES, 1)
        dr = cc2.radio("方向", ["遞增", "遞減"])
        ct = cc3.number_input("數量", 1, 200, 10)
        if st.form_submit_button("執行登錄"):
            plist = []
            cur = sp
            for _ in range(int(ct)):
                if 1 <= cur <= TOTAL_PILES:
                    plist.append(f"{cur}")
                cur = cur + step if dr == "遞增" else cur - step
            process_and_save(plist)

with t2:
    with st.form("m"):
        raw = st.text_input("區間 (支援直接輸入文字與數字，例如 A1-A3, BC1-BC6, 27-30)")
        if st.form_submit_button("執行登錄"):
            process_and_save(parse_range_generic(raw, TOTAL_PILES))

st.divider()

st.markdown("### 🔍 歷史狀態查詢 (時光機)")
df_history_plot = df_history_full.copy()
pdf_report_date = datetime.date.today()
selected_query_date = "最新進度 (預設)"

if not df_history_full.empty:
    unique_dates = sorted(df_history_full['施工日期'].dropna().unique(), reverse=True)
    unique_dates.insert(0, "最新進度 (預設)")
    selected_query_date = st.selectbox("選擇日期 (下方數據、地圖與報表將同步時光倒流至該日狀態)：", unique_dates, key="query_date_picker")

    if selected_query_date != "最新進度 (預設)":
        target_dt = pd.to_datetime(selected_query_date)
        pdf_report_date = target_dt.date()
        df_history_plot = df_history_full[df_history_full['施工日期_DT'] <= target_dt].copy()

        df_date_filtered = df_history_full[df_history_full['施工日期'] == selected_query_date]
        q_a = len(df_date_filtered[df_date_filtered['機台'].astype(str).str.upper().str.contains('A')])
        q_b = len(df_date_filtered[df_date_filtered['機台'].astype(str).str.upper().str.contains('B')])
        st.info(f"📅 【{selected_query_date} 當日施作明細】 共 {len(df_date_filtered)} 支 (A機：{q_a} 支，B機：{q_b} 支)")
        with st.expander("點擊展開當日打設清單"):
            st.dataframe(df_date_filtered[['樁號', '機台', '施作順序', '施工日期']], use_container_width=True)

total_done_auto = len(df_history_plot)
total_perc = (total_done_auto / TOTAL_PILES) * 100 if TOTAL_PILES > 0 else 0
today_done_auto_a = today_done_auto_b = cum_done_a = cum_done_b = this_week_done_a = this_week_done_b = 0
week_start_str = ""
today_state_key = ""

if not df_history_plot.empty:
    latest_dt = df_history_plot['施工日期_DT'].max()
    today_data = df_history_plot[df_history_plot['施工日期_DT'] == latest_dt]
    today_done_auto_a = len(today_data[today_data['機台'].astype(str).str.upper().str.contains('A')])
    today_done_auto_b = len(today_data[today_data['機台'].astype(str).str.upper().str.contains('B')])

    cum_done_a = len(df_history_plot[df_history_plot['機台'].astype(str).str.upper().str.contains('A')])
    cum_done_b = len(df_history_plot[df_history_plot['機台'].astype(str).str.upper().str.contains('B')])

    today_state_key = latest_dt.strftime('%m/%d')
    monday = latest_dt - pd.Timedelta(days=latest_dt.weekday())
    this_week_data = df_history_plot[df_history_plot['施工日期_DT'] >= monday]

    if not this_week_data.empty:
        earliest_this_week = this_week_data['施工日期_DT'].min()
        week_start_str = f"{earliest_this_week.year-1911}/{earliest_this_week.month:02d}/{earliest_this_week.day:02d}"
        this_week_done_a = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('A')])
        this_week_done_b = len(this_week_data[this_week_data['機台'].astype(str).str.upper().str.contains('B')])
    else:
        week_start_str = f"{monday.year-1911}/{monday.month:02d}/{monday.day:02d}"

def process_status_logic(df_hist, df_b):
    if df_b is None or df_b.empty:
        return pd.DataFrame(columns=['X', 'Y', '樁號', '狀態', '標籤', '純順序', 'is_horizontal'])

    plot_df = df_b[['樁號', 'X', 'Y', '數字']].copy().sort_values('數字').reset_index(drop=True)
    dx = plot_df['X'].diff().bfill(); dy = plot_df['Y'].diff().bfill()
    dx_fwd = (plot_df['X'].shift(-1) - plot_df['X']).ffill(); dy_fwd = (plot_df['Y'].shift(-1) - plot_df['Y']).ffill()
    plot_df['is_horizontal'] = (dx + dx_fwd).abs() >= (dy + dy_fwd).abs()

    if df_hist.empty:
        plot_df['狀態'] = '未完成'; plot_df['標籤'] = plot_df['樁號']; plot_df['純順序'] = ""
        return plot_df

    hist = df_hist.copy()
    hist['標籤'] = hist.apply(lambda r: f"{r['樁號']}({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)
    hist['純順序'] = hist.apply(lambda r: f"({str(r.get('機台','A'))[0]}{int(r.get('施作順序',0))})", axis=1)

    max_date = hist['施工日期_DT'].max(); monday_dt = max_date - pd.Timedelta(days=max_date.weekday())
    hist['狀態'] = hist['施工日期_DT'].apply(lambda dt: '未完成' if pd.isna(dt) else ('[已完成]' if dt < monday_dt else dt.strftime('%m/%d')))

    plot_df = plot_df.merge(hist[['樁號', '狀態', '標籤', '純順序']], on='樁號', how='left')
    plot_df['狀態'] = plot_df['狀態'].fillna('未完成'); plot_df['標籤'] = plot_df['標籤'].fillna(plot_df['樁號']); plot_df['純順序'] = plot_df['純順序'].fillna("")
    return plot_df

df_p = process_status_logic(df_history_plot, df_base)

def get_local_stats(sel_list, p_df):
    if not sel_list or p_df.empty:
        return 0, 0
    sub = p_df[p_df['樁號'].isin(sel_list)]
    total = len(sub)
    done = len(sub[sub['狀態'] != '未完成'])
    return done, total

local_a_done, local_a_total = get_local_stats(st.session_state.sel_a, df_p)
local_b_done, local_b_total = get_local_stats(st.session_state.sel_b, df_p)

st.divider()

if not df_p.empty:
    fig_web = px.scatter(df_p, x='X', y='Y', text='標籤', color='狀態', color_discrete_map={'未完成': '#696969', '[已完成]': '#FFB6C1'}, custom_data=['樁號'])
    fig_web.update_traces(selector=dict(name='未完成'), marker=dict(symbol='circle-open', size=16, line=dict(width=2, color='#A9A9A9')), textposition='top right')
    fig_web.update_traces(selector=lambda t: t.name != '未完成', marker=dict(symbol='circle', size=16, line=dict(width=1, color='white')), textposition='top right')
else:
    fig_web = go.Figure()

for loop_pts in boundary_loops:
    if not loop_pts:
        continue
    xs = [p[0] for p in loop_pts]; ys = [p[1] for p in loop_pts]
    if (xs[0], ys[0]) != (xs[-1], ys[-1]):
        xs.append(xs[0]); ys.append(ys[0])
    fig_web.add_trace(
        go.Scatter(x=xs, y=ys, mode='lines',
                   line=dict(color='#D3D3D3', width=2),
                   name='開挖邊界', hoverinfo='skip', showlegend=False)
    )

fig_web.update_layout(xaxis_visible=False, yaxis=dict(scaleanchor="x", scaleratio=1, visible=False), height=900, plot_bgcolor='white', dragmode='pan')

st.subheader("🗺️ 網頁選取區 (框選或輸入以擷取局部圖)")
try:
    selection_event = st.plotly_chart(fig_web, use_container_width=True, config={'scrollZoom': True}, on_select="rerun", selection_mode=('box', 'lasso'))
    selected_piles = [pt["customdata"][0] for pt in selection_event["selection"]["points"]] if selection_event and "selection" in selection_event and selection_event["selection"]["points"] else []
except Exception:
    selected_piles = []

if selected_piles:
    st.success(f"🎯 畫面上滑鼠目前已選取： **{len(selected_piles)}** 支樁位")
else:
    st.caption("💡 提示：請在地圖上方拉框選取，或直接使用下方文字輸入範圍。")

st.markdown("#### ⚙️ 分配 PDF 局部截圖範圍")
c_btn1, c_btn2, c_btn3 = st.columns([1.5, 2, 1])

with c_btn1:
    st.markdown("**👉 方式一：將【滑鼠框選】的範圍分配給**")
    cb1, cb2 = st.columns(2)
    if cb1.button("📌 A機 (框選)"): st.session_state.sel_a = selected_piles; st.rerun()
    if cb2.button("📌 B機 (框選)"): st.session_state.sel_b = selected_piles; st.rerun()

with c_btn2:
    st.markdown("**👉 方式二：將【文字輸入】的範圍分配給**")
    manual_raw = st.text_input("輸入樁號區間 (如: 27-30, BC1-BC6, A1)", label_visibility="collapsed")
    cb3, cb4 = st.columns(2)
    if cb3.button("📌 A機 (輸入)"): st.session_state.sel_a = parse_range_generic(manual_raw, TOTAL_PILES); st.rerun()
    if cb4.button("📌 B機 (輸入)"): st.session_state.sel_b = parse_range_generic(manual_raw, TOTAL_PILES); st.rerun()

with c_btn3:
    st.markdown("**🗑️ 重新設定**")
    if st.button("清除所有截圖", use_container_width=True): st.session_state.sel_a = []; st.session_state.sel_b = []; st.rerun()

st.info(f"當前 PDF 暫存狀態：A機截圖區包含 {len(st.session_state.sel_a)} 支樁 | B機截圖區包含 {len(st.session_state.sel_b)} 支樁")

if not df_history_plot.empty or not df_p.empty:
    st.sidebar.markdown("### 📄 PDF 報表文字內容")
    st.sidebar.text_input("右側主標題", key="pdf_loc_note_right")
    st.sidebar.text_input("左側副標題", key="pdf_loc_note_left")
    st.sidebar.number_input("本週預計完成 (支)", key="pdf_week_est", step=1)
    show_seq = st.sidebar.checkbox("🔢 PDF 圖上顯示施作順序 (機台+順序號)", value=True)

    st.sidebar.markdown("### 🎛️ PDF 圖表幾何微調")
    with st.sidebar.form("geom"):
        fig_scale = st.slider("排樁間距拉開倍率", 1.0, 5.0, s['fig_scale'], 0.1)
        marker_size = st.slider("圓圈大小", 50, 400, s['marker_size'], 10)
        lbl_fontsize = st.slider("樁號文字大小", 8, 40, s['lbl_fontsize'], 1)
        text_offset = st.slider("文字離圓圈距離", 5, 60, s['text_offset'], 1)
        st.form_submit_button("🔄 套用幾何設定")

    st.sidebar.markdown("### 📐 PDF 文字與截圖位置微調")
    with st.sidebar.form("layout"):
        pos_title_y = st.slider("大標題高度 (Y)", 0.0, 1.0, s['pos_title_y'], 0.01)
        pos_info_x = st.slider("資訊區左右 (X)", 0.0, 1.0, s['pos_info_x'], 0.01)
        pos_info_y = st.slider("資訊區高度 (Y)", 0.0, 1.0, s['pos_info_y'], 0.01)
        pos_loc_x = st.slider("右側標題 (X)", 0.0, 1.0, s['pos_loc_x'], 0.01)
        pos_loc_y = st.slider("右側標題 (Y)", 0.0, 1.0, s['pos_loc_y'], 0.01)
        pos_loc_x_left = st.slider("左側標題 (X)", 0.0, 1.0, s['pos_loc_x_left'], 0.01)
        pos_loc_y_left = st.slider("左側標題 (Y)", 0.0, 1.0, s['pos_loc_y_left'], 0.01)
        pos_leg_x = st.slider("圖例左右 (X)", -1.0, 1.5, s['pos_leg_x'], 0.01)
        pos_leg_y = st.slider("圖例高度 (Y)", -1.0, 1.5, s['pos_leg_y'], 0.01)

        st.markdown("#### 局部預覽圖位置微調")
        pos_img_a_x = st.slider("A機圖 左右 (X)", 0.0, 1.0, s.get('pos_img_a_x', 0.35), 0.01)
        pos_img_a_y = st.slider("A機圖 高度 (Y)", 0.0, 1.0, s.get('pos_img_a_y', 0.10), 0.01)
        pos_img_a_w = st.slider("A機圖 寬度 (W)", 0.1, 1.0, s.get('pos_img_a_w', 0.30), 0.01)

        pos_img_b_x = st.slider("B機圖 左右 (X)", 0.0, 1.0, s.get('pos_img_b_x', 0.68), 0.01)
        pos_img_b_y = st.slider("B機圖 高度 (Y)", 0.0, 1.0, s.get('pos_img_b_y', 0.10), 0.01)
        pos_img_b_w = st.slider("B機圖 寬度 (W)", 0.1, 1.0, s.get('pos_img_b_w', 0.30), 0.01)

        st.form_submit_button("🔄 套用排版與圖位設定")

    if st.sidebar.button("💾 記憶當前排版與標題 (永久儲存)"):
        new_s = {
            "pdf_loc_note_right": st.session_state.pdf_loc_note_right,
            "pdf_loc_note_left": st.session_state.pdf_loc_note_left,
            "pdf_week_est": st.session_state.pdf_week_est,
            "fig_scale": fig_scale, "marker_size": marker_size, "lbl_fontsize": lbl_fontsize, "text_offset": text_offset,
            "pos_title_y": pos_title_y, "pos_info_x": pos_info_x, "pos_info_y": pos_info_y,
            "pos_loc_x": pos_loc_x, "pos_loc_y": pos_loc_y, "pos_loc_x_left": pos_loc_x_left, "pos_loc_y_left": pos_loc_y_left,
            "pos_leg_x": pos_leg_x, "pos_leg_y": pos_leg_y,
            "pos_img_a_x": pos_img_a_x, "pos_img_a_y": pos_img_a_y, "pos_img_a_w": pos_img_a_w,
            "pos_img_b_x": pos_img_b_x, "pos_img_b_y": pos_img_b_y, "pos_img_b_w": pos_img_b_w
        }
        if demo_mode:
            st.sidebar.warning("👀 沙盒試用：設定僅變更於當前瀏覽器，未寫入雲端檔案。")
            s.update(new_s)
        else:
            save_settings(ss, SETTINGS_WS_NAME, new_s)
            st.session_state.ui_settings = new_s
            st.sidebar.success("✅ 設定已寫入雲端永久記憶")

    def draw_pdf_axis(ax, target_df, global_df, scale_factor=1.0, is_main=False, show_seq=True):
        label_texts = []
        label_points_x = []
        label_points_y = []

        if target_df.empty and not boundary_loops:
            ax.axis('off')
            return

        if not target_df.empty and len(target_df) < len(global_df):
            x_min, x_max = target_df['X'].min(), target_df['X'].max()
            y_min, y_max = target_df['Y'].min(), target_df['Y'].max()
            pad_x = max((x_max - x_min) * 0.2, 500)
            pad_y = max((y_max - y_min) * 0.2, 500)

            for i, loop_pts in enumerate(boundary_loops):
                if not loop_pts:
                    continue
                xs = [p[0] for p in loop_pts]; ys = [p[1] for p in loop_pts]
                if (xs[0], ys[0]) != (xs[-1], ys[-1]):
                    xs.append(xs[0]); ys.append(ys[0])
                loop_x = np.array(xs); loop_y = np.array(ys)
                mask = (loop_x >= x_min - pad_x) & (loop_x <= x_max + pad_x) & (loop_y >= y_min - pad_y) & (loop_y <= y_max + pad_y)
                if mask.any():
                    lbl = '開挖邊界' if is_main and i == 0 else None
                    ax.plot(loop_x, loop_y, color='#E0E0E0', linewidth=2, zorder=1, label=lbl)

            ax.set_xlim(x_min - pad_x, x_max + pad_x)
            ax.set_ylim(y_min - pad_y, y_max + pad_y)
        else:
            for i, loop_pts in enumerate(boundary_loops):
                if not loop_pts:
                    continue
                xs = [p[0] for p in loop_pts]; ys = [p[1] for p in loop_pts]
                if (xs[0], ys[0]) != (xs[-1], ys[-1]):
                    xs.append(xs[0]); ys.append(ys[0])
                lbl = '開挖邊界' if is_main and i == 0 else None
                ax.plot(xs, ys, color='#E0E0E0', linewidth=2, zorder=1, label=lbl)

        states = ['未完成', '[已完成]'] + sorted([st_ for st_ in global_df['狀態'].unique() if st_ not in ['未完成', '[已完成]']])
        colors = {'未完成': '#808080', '[已完成]': '#FFB6C1'}
        pal = px.colors.qualitative.Plotly
        color_idx = 0
        for s_glob in states:
            if s_glob not in colors:
                colors[s_glob] = pal[color_idx % len(pal)]
                color_idx += 1

        msize = marker_size * scale_factor
        fsize = lbl_fontsize * scale_factor
        offset = text_offset * scale_factor

        for state in states:
            sub = target_df[target_df['狀態'] == state]
            c = colors[state]

            if state == '未完成':
                legend_label = "未完成" if is_main else None
                if not sub.empty:
                    ax.scatter(sub['X'], sub['Y'], facecolors='none', edgecolors=c, s=msize, lw=1.5, zorder=2, label=legend_label)
                elif is_main:
                    ax.scatter([], [], facecolors='none', edgecolors=c, s=msize, lw=1.5, zorder=2, label=legend_label)
            else:
                legend_label = (f"{state} 樁號 ○ 施作順序" if show_seq else f"{state} 樁號") if is_main else None
                if not sub.empty:
                    ax.scatter(sub['X'], sub['Y'], color=c, s=msize, zorder=3, label=legend_label)
                    if state == today_state_key:
                        for _, row in sub.iterrows():
                            p = row['樁號']; s_txt = row['純順序']
                            combo = f"{p}\n{s_txt}" if (show_seq and s_txt) else p
                            if ADJUSTTEXT_READY:
                                t = ax.text(row['X'], row['Y'], combo, fontsize=fsize,
                                            fontweight='bold', color='black', ha='center', va='center', zorder=5)
                            else:
                                is_h = row['is_horizontal']
                                if is_h:
                                    t = ax.text(row['X'], row['Y'] + offset, combo, fontsize=fsize,
                                                fontweight='bold', color='black', ha='center', va='bottom', zorder=5)
                                else:
                                    t = ax.text(row['X'] - offset, row['Y'], combo, fontsize=fsize,
                                                fontweight='bold', color='black', ha='right', va='center', zorder=5)
                            label_texts.append(t)
                            label_points_x.append(row['X'])
                            label_points_y.append(row['Y'])
                elif is_main:
                    ax.scatter([], [], color=c, s=msize, zorder=3, label=legend_label)

        if ADJUSTTEXT_READY and label_texts:
            adjust_text(
                label_texts,
                x=label_points_x, y=label_points_y,
                ax=ax,
                arrowprops=dict(arrowstyle='-', color='gray', lw=0.7, alpha=0.8),
                expand_text=(1.5, 1.8),
                expand_points=(2.2, 2.5),
                force_text=(0.6, 0.9),
                force_points=(0.8, 1.0),
                lim=2000,
            )

        ax.margins(0.1)
        ax.set_aspect('equal', adjustable='datalim')
        ax.axis('off')

    if MATPLOTLIB_READY:
        @st.cache_resource
        def setup_chinese_font():
            import urllib.request
            import matplotlib.font_manager as fm
            font_path = 'NotoSansCJKtc-Regular.otf'
            if not os.path.exists(font_path):
                try:
                    url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
                    urllib.request.urlretrieve(url, font_path)
                except Exception:
                    pass
            if os.path.exists(font_path):
                fm.fontManager.addfont(font_path)
                return fm.FontProperties(fname=font_path).get_name()
            return None

        def create_pdf_figure():
            font_name = setup_chinese_font()
            if font_name: plt.rcParams['font.family'] = font_name
            fig = plt.figure(figsize=(24 * fig_scale, 16 * fig_scale))
            has_a, has_b = bool(st.session_state.sel_a), bool(st.session_state.sel_b)

            if not (has_a or has_b):
                ax = fig.add_axes([0.45, 0.1, 0.5, 0.75])
                draw_pdf_axis(ax, df_p, df_p, 1.0, True, show_seq)
                ax.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28 * fig_scale, markerscale=1.5)
            else:
                if has_a and has_b:
                    ax_a = fig.add_axes([pos_img_a_x, pos_img_a_y, pos_img_a_w, 0.75])
                    draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], df_p, 1.0, True, show_seq)
                    ax_a.set_title("A機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05)
                    ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)

                    ax_b = fig.add_axes([pos_img_b_x, pos_img_b_y, pos_img_b_w, 0.75])
                    draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], df_p, 1.0, False, show_seq)
                    ax_b.set_title("B機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05)
                elif has_a:
                    ax_a = fig.add_axes([pos_img_a_x, pos_img_a_y, pos_img_a_w, 0.75])
                    draw_pdf_axis(ax_a, df_p[df_p['樁號'].isin(st.session_state.sel_a)], df_p, 1.0, True, show_seq)
                    ax_a.set_title("A機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05)
                    ax_a.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)
                elif has_b:
                    ax_b = fig.add_axes([pos_img_b_x, pos_img_b_y, pos_img_b_w, 0.75])
                    draw_pdf_axis(ax_b, df_p[df_p['樁號'].isin(st.session_state.sel_b)], df_p, 1.0, True, show_seq)
                    ax_b.set_title("B機作業區", fontsize=40*fig_scale, fontweight='bold', y=-0.05)
                    ax_b.legend(loc='lower left', bbox_to_anchor=(pos_leg_x, pos_leg_y), fontsize=28*fig_scale, markerscale=1.5)

            roc_y = pdf_report_date.year - 1911
            pdf_title_date = f"{roc_y}/{pdf_report_date.month:02d}/{pdf_report_date.day:02d}"

            a_pct_str = f" ({(local_a_done/local_a_total)*100:.2f}%)" if local_a_total > 0 else ""
            b_pct_str = f" ({(local_b_done/local_b_total)*100:.2f}%)" if local_b_total > 0 else ""

            info_lines = [
                f"本週預計完成 {st.session_state.pdf_week_est} 支",
                f"{week_start_str}至{pdf_title_date}",
                f"本週累積 A機:{this_week_done_a}支 B機:{this_week_done_b}支",
                f"該日完成 A機:{today_done_auto_a}支 B機:{today_done_auto_b}支",
                f"選取區 A機:{local_a_done}/{local_a_total}{a_pct_str}",
                f"    B機:{local_b_done}/{local_b_total}{b_pct_str}",
                f"總累積完成 {total_done_auto} 支 ({total_done_auto}/{TOTAL_PILES}, {total_perc:.2f}%)",
                f"各別累積 A機:{cum_done_a}支 B機:{cum_done_b}支"
            ]
            fig.text(0.05, pos_title_y, f"{pdf_title_date} 施作進度回報", fontsize=50 * fig_scale, fontweight='bold')
            fig.text(pos_info_x, pos_info_y, "\n".join(info_lines), fontsize=35 * fig_scale, linespacing=1.6, va='top')
            fig.text(pos_loc_x, pos_loc_y, st.session_state.pdf_loc_note_right, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            fig.text(pos_loc_x_left, pos_loc_y_left, st.session_state.pdf_loc_note_left, fontsize=55 * fig_scale, fontweight='bold', ha='center')
            return fig

        pdf_fig = create_pdf_figure(); st.divider()
        if ADJUSTTEXT_READY:
            st.caption("✅ 文字防重疊模組 (adjustText) 已啟用")
        else:
            st.caption("⚠️ 文字防重疊模組 (adjustText) 未安裝成功，目前使用備援手動偏移。請確認 requirements.txt 內有 adjustText 並重新 Reboot App。")
        st.pyplot(pdf_fig)
        buf = io.BytesIO(); pdf_fig.savefig(buf, format='pdf', bbox_inches='tight'); plt.close(pdf_fig)
        st.sidebar.markdown("### 📥 下載區")
        has_local_download = bool(st.session_state.sel_a) or bool(st.session_state.sel_b)
        pdf_btn_text = "🔴 匯出 PDF 報表 (局部圖)" if has_local_download else "🔴 匯出 PDF 報表 (全區圖)"
        st.sidebar.download_button(pdf_btn_text, buf.getvalue(), f"{site_id}_Plan_{pdf_report_date}.pdf", type="primary")

    def xl_gen(h_df, p_df):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
            h_df.to_excel(wr, sheet_name='施工明細', index=False)
            wb = wr.book; ws = wb.add_worksheet('全區進度圖'); ch = wb.add_chart({'type': 'scatter'})
            col = 10; states = ['未完成', '[已完成]'] + sorted([st_ for st_ in p_df['狀態'].unique() if st_ not in ['未完成', '[已完成]']])
            colors = {'未完成': '#696969', '[已完成]': '#FFB6C1'}; pal = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']; ci = 0
            for st_ in states:
                sub = p_df[p_df['狀態'] == st_].reset_index(drop=True)
                if sub.empty:
                    continue
                sub[['X', 'Y', '標籤']].to_excel(wr, sheet_name='全區進度圖', startcol=col, index=False)
                mc = colors.get(st_, pal[ci % len(pal)])
                if st_ not in colors: ci += 1
                sd = {'name': st_, 'categories': ['全區進度圖', 1, col, len(sub), col], 'values': ['全區進度圖', 1, col+1, len(sub), col+1], 'marker': {'type': 'circle', 'size': 6, 'fill': {'color': mc}, 'border': {'color': mc}}}
                if st_ == '未完成': sd['marker']['fill'] = {'none': True}
                if st_ != '未完成': sd['data_labels'] = {'custom': [{'value': f'=全區進度圖!${xlsxwriter.utility.xl_col_to_name(col+2)}${ri+2}'} for ri in range(len(sub))], 'position': 'above', 'font': {'size': 8}}
                ch.add_series(sd); col += 4
            ch.set_x_axis({'visible': False}); ch.set_y_axis({'visible': False}); ws.insert_chart('B2', ch)
        return out.getvalue()

    st.sidebar.download_button("🟢 匯出 Excel (全區報表)", xl_gen(df_history_plot, df_p), f"{site_id}_Report_{pdf_report_date}.xlsx")

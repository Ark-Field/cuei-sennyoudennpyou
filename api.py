import io
import os
import json
import requests
import pandas as pd
from datetime import datetime, timedelta, date
import streamlit as st
import copy
import openpyxl

# ReportLab関連（PDF出力用）
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# 0. 日本語フォント（NotoSansJP）の登録関数
# ==========================================
def register_jp_font():
    font_paths = [
        "NotoSansJP-Regular.ttf",
        "fonts/NotoSansJP-Regular.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansJP-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansJP-Regular.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("NotoSansJP", path))
                return "NotoSansJP"
            except Exception:
                pass
    return "Helvetica"

JP_FONT = register_jp_font()

# ==========================================
# 1. ページ基本設定
# ==========================================
st.set_page_config(page_title="製本・納品管理システム", layout="wide")
st.title("📊 製本・納品・専用伝票管理システム")

# ==========================================
# 2. Secrets設定
# ==========================================
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID = 34  # @pocket対象アプリID

# ==========================================
# 3. APIデータ取得関数
# ==========================================
def fetch_pocket_records(start_date=None, end_date=None, target_order_no=None, target_publisher=None):
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}
    
    params = {}
    
    if target_order_no:
        params["query"] = f'field-2 = "{target_order_no}"'
    else:
        query_parts = []
        if start_date and end_date:
            start_str = start_date.strftime("%Y-%m-%d")
            end_str = end_date.strftime("%Y-%m-%d")
            query_parts.append(f'field-1 >= "{start_str}" and field-1 <= "{end_str}"')
        if target_publisher:
            query_parts.append(f'field-8 like "{target_publisher}"')
        if query_parts:
            params["query"] = " and ".join(query_parts)
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=20)
        if response.status_code != 200 and "query" in params:
            response = requests.get(url, headers=headers, timeout=20)
        
        if response.status_code == 200:
            result = response.json()
            raw_items = result.get("records", result.get("Data", result.get("Records", [])))
            
            parsed_records = []
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                rec_body = item.get("record", {})
                if not isinstance(rec_body, dict):
                    rec_body = {}
                
                merged_body = {**item, **rec_body}
                merged_body["_recordId"] = item.get("Id", item.get("recordId", item.get("ResultId", rec_body.get("Id", ""))))
                parsed_records.append(merged_body)
            
            return parsed_records
        else:
            st.error(f"❌ APIエラー (ステータスコード {response.status_code}): {response.text}")
    except Exception as e:
        st.error(f"🚨 通信例外エラー: {e}")
    
    return []

def parse_date_safe(val):
    if not val or str(val).lower() in ["nan", "none", ""]:
        return None
    try:
        return datetime.strptime(str(val)[:10].replace("/", "-"), "%Y-%m-%d").date()
    except ValueError:
        return None

def safe_float(val):
    try:
        if val is None or str(val).strip() == "" or str(val).lower() in ["nan", "none"]:
            return 0.0
        return float(str(val).replace(",", ""))
    except ValueError:
        return 0.0

# ==========================================
# 4. 初期データロード
# ==========================================
if "api_records" not in st.session_state:
    default_end = date.today()
    default_start = default_end - timedelta(days=60)
    st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end)

api_records = st.session_state.get("api_records", [])

# ==========================================
# 5. タブ構成の定義
# ==========================================
tab1, tab2 = st.tabs([
    "📋 タブ1: 受注一覧表 ＆ データ更新", 
    "📦 タブ2: 納品予定表 ＆ 専用伝票発行（顧客×便別）"
])

# ==========================================
# フラットデータ（明細行展開）の構築
# ==========================================
def build_flat_data(records):
    flat = []
    for record_idx, r in enumerate(records):
        h_ord_no = str(r.get("field-2", ""))
        h_ord_date = parse_date_safe(r.get("field-1"))
        h_pub = str(r.get("field-8", ""))
        h_title = str(r.get("field-13", ""))
        h_format = str(r.get("field-23", ""))
        h_isbn = str(r.get("field-93", ""))
        
        h_print_qty = str(r.get("field-3", ""))
        h_qty = str(r.get("field-4", ""))
        h_tanka = str(r.get("field-5", ""))
        h_price = str(r.get("field-20", ""))
        h_shukai = str(r.get("field-29", ""))
        h_deposit = str(r.get("field-81", ""))
        
        if h_tanka.lower() in ["nan", "none", ""]:
            h_tanka = "0"
        if h_shukai.lower() in ["nan", "none", ""]:
            h_shukai = ""

        sub_90 = r.get("field-90", [])
        sub_88 = r.get("field-88", [])
        sub_89 = r.get("field-89", [])
        
        max_len = max(len(sub_90) if isinstance(sub_90, list) else 0,
                      len(sub_88) if isinstance(sub_88, list) else 0,
                      len(sub_89) if isinstance(sub_89, list) else 0, 1)
        
        for i in range(max_len):
            item_90 = sub_90[i] if isinstance(sub_90, list) and len(sub_90) > i else {}
            sub_id_90 = item_90.get("recordSubId", item_90.get("Id", ""))
            rec_90 = item_90.get("record", item_90) if isinstance(item_90, dict) else {}
            
            item_88 = sub_88[i] if isinstance(sub_88, list) and len(sub_88) > i else {}
            sub_id_88 = item_88.get("recordSubId", item_88.get("Id", ""))
            rec_88 = item_88.get("record", item_88) if isinstance(item_88, dict) else {}
            
            haisha_val = str(rec_88.get("field-88_10", ""))
            if not haisha_val or haisha_val.lower() in ["nan", "none", ""]:
                haisha_val = ""

            flat.append({
                "更新": False,
                "_record_idx": record_idx,
                "_sub_idx": i,
                "_id_90": sub_id_90,
                "_id_88": sub_id_88,
                "受注番号": h_ord_no,
                "受注日": h_ord_date,
                "出版社名": h_pub,
                "書名": h_title,
                "判型": h_format,
                "ISBN": h_isbn,
                "印刷部数": h_print_qty,
                "受注数量": h_qty,
                "定価": h_price,
                "預かり在庫": h_deposit,
                "予定日": parse_date_safe(rec_90.get("field-90_8")),
                "納品日": parse_date_safe(rec_88.get("field-88_2")),
                "納品部数": str(rec_88.get("field-88_4", "")),
                "PL数": str(rec_88.get("field-88_7", "")),
                "見本": str(rec_88.get("field-88_8", "")),
                "刷回": h_shukai,
                "単価": h_tanka,
                "摘要": str(rec_88.get("field-88_5", "")),
                "在庫納品": bool(rec_88.get("field-88_6", False)),
                "専用伝票番号": str(rec_88.get("field-88_9", "")),
                "配車区分": haisha_val,
            })
    return flat

df_flat = pd.DataFrame(build_flat_data(api_records))

# ==========================================
# タブ1：受注一覧表 ＆ データ更新
# ==========================================
with tab1:
    st.header("📋 受注一覧表 ＆ データ更新エディタ")
    
    with st.expander("🔍 検索条件の変更・絞り込み", expanded=True):
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            default_end = date.today()
            default_start = default_end - timedelta(days=60)
            t1_start = st.date_input("受注日（開始）", value=default_start, key="t1_s")
        with col_s2:
            t1_end = st.date_input("受注日（終了）", value=default_end, key="t1_e")
        with col_s3:
            publisher_input = st.text_input("📖 出版社名で絞り込み", value="", key="pub_ord")
        
        if st.button("🔄 検索実行", type="primary", key="btn_s1"):
            with st.spinner("@ポケットからデータを取得中..."):
                st.session_state["api_records"] = fetch_pocket_records(
                    start_date=t1_start, 
                    end_date=t1_end, 
                    target_publisher=publisher_input if publisher_input else None
                )
            st.rerun()

    if not df_flat.empty:
        edited_flat_df = st.data_editor(
            df_flat,
            num_rows="fixed",
            key="grid_flat_main",
            use_container_width=True,
            column_config={
                "更新": st.column_config.CheckboxColumn("更新", required=True, help="チェックした行を@pocketに上書き保存"),
                "_record_idx": None, "_sub_idx": None, "_id_90": None, "_id_88": None,
                "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
                "受注日": st.column_config.DateColumn("受注日", format="YYYY-MM-DD", disabled=True),
                "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
                "書名": st.column_config.TextColumn("書名", disabled=True),
                "判型": st.column_config.TextColumn("判型(23)"),
                "ISBN": st.column_config.TextColumn("ISBN", disabled=True),
                "印刷部数": st.column_config.TextColumn("印刷部数(3)"),
                "受注数量": st.column_config.TextColumn("受注数量(4)"),
                "定価": st.column_config.TextColumn("定価", disabled=True),
                "預かり在庫": st.column_config.TextColumn("預かり在庫(81)"),
                "予定日": st.column_config.DateColumn("予定日(90_8)", format="YYYY-MM-DD", disabled=True),
                "納品日": st.column_config.DateColumn("納品日", format="YYYY-MM-DD"),
                "納品部数": st.column_config.TextColumn("納品部数"),
                "PL数": st.column_config.TextColumn("PL数"),
                "見本": st.column_config.TextColumn("見本"),
                "刷回": st.column_config.TextColumn("刷回(29)"),
                "単価": st.column_config.TextColumn("単価(5)"),
                "摘要": st.column_config.TextColumn("摘要"),
                "在庫納品": st.column_config.CheckboxColumn("在庫納品"),
                "専用伝票番号": st.column_config.TextColumn("伝票番号"),
                "配車区分": st.column_config.SelectboxColumn("配車区分", options=["", "A便", "B便", "C便", "自社便", "宅配便"]),
            }
        )
        
        st.divider()
        if st.button("💾 チェックを入れたデータをデータベースに一括上書き保存する", type="primary", key="btn_save_checked"):
            checked_rows = edited_flat_df[edited_flat_df["更新"] == True]
            if checked_rows.empty:
                st.warning("⚠️ 保存する行の「更新」チェックボックスにチェックを入れてください。")
            else:
                headers = {
                    "X-At-Pocket-API-Key": POCKET_API_KEY, 
                    "Accept": "application/json", 
                    "Content-Type": "application/json"
                }
                success_count = 0
                
                for r_idx, group_df in checked_rows.groupby("_record_idx"):
                    r_idx = int(r_idx)
                    original_rec = api_records[r_idx]
                    record_id = original_rec.get("_recordId")
                    if not record_id:
                        continue
                    
                    update_url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records/{record_id}"
                    first_row = group_df.iloc[0]
                    record_payload = {
                        "field-1": str(first_row["受注日"]) if pd.notna(first_row["受注日"]) else "",
                        "field-2": str(first_row["受注番号"]),
                        "field-3": str(first_row["印刷部数"]),
                        "field-4": str(first_row["受注数量"]),
                        "field-5": str(first_row["単価"]),
                        "field-20": str(first_row["定価"]),
                        "field-23": str(first_row["判型"]),
                        "field-29": str(first_row["刷回"]),
                        "field-81": str(first_row["預かり在庫"]),
                        "field-93": str(first_row["ISBN"]),
                    }
                    
                    orig_sub_88 = original_rec.get("field-88", [])
                    sub88_list = []
                    for _, row in group_df.iterrows():
                        sub_i = int(row["_sub_idx"])
                        item_88 = copy.deepcopy(orig_sub_88[sub_i]) if (isinstance(orig_sub_88, list) and len(orig_sub_88) > sub_i) else {}
                        rec_dict = item_88.get("record", {}) if isinstance(item_88, dict) else {}
                        
                        rec_dict["field-88_2"] = str(row["納品日"]) if pd.notna(row["納品日"]) else ""
                        rec_dict["field-88_4"] = str(row["納品部数"])
                        rec_dict["field-88_5"] = str(row["摘要"])
                        rec_dict["field-88_6"] = bool(row["在庫納品"])
                        rec_dict["field-88_7"] = str(row["PL数"])
                        rec_dict["field-88_8"] = str(row["見本"])
                        rec_dict["field-88_9"] = str(row["専用伝票番号"])
                        
                        h_val = row["配車区分"]
                        if pd.isna(h_val) or str(h_val).lower() in ["nan", "none", ""]:
                            rec_dict["field-88_10"] = ""
                        else:
                            rec_dict["field-88_10"] = str(h_val)
                        
                        item_88["record"] = rec_dict
                        sub88_list.append(item_88)
                    
                    record_payload["field-88"] = sub88_list
                    payload = {"apiVersion": 1.1, "record": record_payload}
                    
                    try:
                        res = requests.put(update_url, headers=headers, json=payload, timeout=15)
                        if res.status_code in [200, 204]:
                            success_count += 1
                        else:
                            st.error(f"❌ 書き戻しエラー (ID: {record_id}): {res.text}")
                    except Exception as ex:
                        st.error(f"🚨 例外発生: {ex}")
                
                if success_count > 0:
                    st.success(f"✅ {success_count} 件のレコードを正常に上書き保存しました！")
                    default_end = date.today()
                    default_start = default_end - timedelta(days=60)
                    st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end)
                    st.rerun()
    else:
        st.info("ℹ️ 表示するデータがありません。")

# ==========================================
# タブ2：納品予定表 ＆ 専用伝票発行（顧客×便別）
# ==========================================
with tab2:
    st.header("📦 納品予定表 ＆ 専用伝票発行（顧客別 × 便別）")
    
    col_t2_1, col_t2_2 = st.columns(2)
    with col_t2_1:
        target_delivery_date = st.date_input("対象予定日を指定 (field-90_8)", value=date.today(), key="target_deliv_date")
    with col_t2_2:
        auto_copy_delivery = st.checkbox("💡 予定日を確定納品日(88_2)へ自動転記する", value=False)
    
    target_date_str = target_delivery_date.strftime("%Y-%m-%d")
    
    if not df_flat.empty:
        df_t2_raw = df_flat[df_flat["予定日"].astype(str) == target_date_str].copy()
    else:
        df_t2_raw = pd.DataFrame()
    
    if not df_t2_raw.empty:
        if auto_copy_delivery:
            for idx, row in df_t2_raw.iterrows():
                if pd.isna(row["納品日"]) or str(row["納品日"]) in ["NaT", "", "None"]:
                    df_t2_raw.at[idx, "納品日"] = target_delivery_date
                    
        st.success(f"✨ 予定日【 {target_date_str} 】の明細データが **{len(df_t2_raw)} 件** ヒットしました！下の表で直接入力・編集できます。")
        
        edited_t2_df = st.data_editor(
            df_t2_raw,
            num_rows="fixed",
            key="grid_t2_edit",
            use_container_width=True,
            column_config={
                "更新": st.column_config.CheckboxColumn("更新", required=True, help="この行の変更を@pocketに書き戻す"),
                "_record_idx": None, "_sub_idx": None, "_id_90": None, "_id_88": None,
                "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
                "受注日": st.column_config.DateColumn("受注日", format="YYYY-MM-DD", disabled=True),
                "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
                "書名": st.column_config.TextColumn("書名", disabled=True),
                "判型": st.column_config.TextColumn("判型"),
                "ISBN": st.column_config.TextColumn("ISBN", disabled=True),
                "印刷部数": st.column_config.TextColumn("印刷部数"),
                "受注数量": st.column_config.TextColumn("受注数量"),
                "定価": st.column_config.TextColumn("定価", disabled=True),
                "預かり在庫": st.column_config.TextColumn("預かり在庫"),
                "予定日": st.column_config.DateColumn("予定日", format="YYYY-MM-DD", disabled=True),
                "納品日": st.column_config.DateColumn("納品日", format="YYYY-MM-DD"),
                "納品部数": st.column_config.TextColumn("納品部数"),
                "PL数": st.column_config.TextColumn("PL数"),
                "見本": st.column_config.TextColumn("見本"),
                "刷回": st.column_config.TextColumn("刷回(29)"),
                "単価": st.column_config.TextColumn("単価(5)"),
                "配車区分": st.column_config.SelectboxColumn("配車区分", options=["", "A便", "B便", "C便", "自社便", "宅配便"]),
                "摘要": st.column_config.TextColumn("摘要"),
                "在庫納品": st.column_config.CheckboxColumn("在庫納品"),
                "専用伝票番号": st.column_config.TextColumn("伝票番号"),
            }
        )
        
        # タブ2専用の書き戻しボタン
        if st.button("💾 タブ2の変更内容を@pocketに書き戻す", type="primary", key="btn_save_t2"):
            checked_t2 = edited_t2_df[edited_t2_df["更新"] == True]
            if checked_t2.empty:
                st.warning("⚠️ 書き戻す行の「更新」チェックボックスにチェックを入れてください。")
            else:
                headers = {
                    "X-At-Pocket-API-Key": POCKET_API_KEY, 
                    "Accept": "application/json", 
                    "Content-Type": "application/json"
                }
                success_count = 0
                
                for r_idx, group_df in checked_t2.groupby("_record_idx"):
                    r_idx = int(r_idx)
                    original_rec = api_records[r_idx]
                    record_id = original_rec.get("_recordId")
                    if not record_id:
                        continue
                    
                    update_url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records/{record_id}"
                    first_row = group_df.iloc[0]
                    record_payload = {
                        "field-1": str(first_row["受注日"]) if pd.notna(first_row["受注日"]) else "",
                        "field-2": str(first_row["受注番号"]),
                        "field-3": str(first_row["印刷部数"]),
                        "field-4": str(first_row["受注数量"]),
                        "field-5": str(first_row["単価"]),
                        "field-20": str(first_row["定価"]),
                        "field-23": str(first_row["判型"]),
                        "field-29": str(first_row["刷回"]),
                        "field-81": str(first_row["預かり在庫"]),
                        "field-93": str(first_row["ISBN"]),
                    }
                    
                    orig_sub_88 = original_rec.get("field-88", [])
                    sub88_list = []
                    for _, row in group_df.iterrows():
                        sub_i = int(row["_sub_idx"])
                        item_88 = copy.deepcopy(orig_sub_88[sub_i]) if (isinstance(orig_sub_88, list) and len(orig_sub_88) > sub_i) else {}
                        rec_dict = item_88.get("record", {}) if isinstance(item_88, dict) else {}
                        
                        rec_dict["field-88_2"] = str(row["納品日"]) if pd.notna(row["納品日"]) else ""
                        rec_dict["field-88_4"] = str(row["納品部数"])
                        rec_dict["field-88_5"] = str(row["摘要"])
                        rec_dict["field-88_6"] = bool(row["在庫納品"])
                        rec_dict["field-88_7"] = str(row["PL数"])
                        rec_dict["field-88_8"] = str(row["見本"])
                        rec_dict["field-88_9"] = str(row["専用伝票番号"])
                        
                        h_val = row["配車区分"]
                        if pd.isna(h_val) or str(h_val).lower() in ["nan", "none", ""]:
                            rec_dict["field-88_10"] = ""
                        else:
                            rec_dict["field-88_10"] = str(h_val)
                        
                        item_88["record"] = rec_dict
                        sub88_list.append(item_88)
                    
                    record_payload["field-88"] = sub88_list
                    payload = {"apiVersion": 1.1, "record": record_payload}
                    
                    try:
                        res = requests.put(update_url, headers=headers, json=payload, timeout=15)
                        if res.status_code in [200, 204]:
                            success_count += 1
                        else:
                            st.error(f"❌ タブ2書き戻しエラー (ID: {record_id}): {res.text}")
                    except Exception as ex:
                        st.error(f"🚨 例外発生: {ex}")
                
                if success_count > 0:
                    st.success(f"✅ {success_count} 件のデータを@pocketへ正常に書き戻しました！")
                    default_end = date.today()
                    default_start = default_end - timedelta(days=60)
                    st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end)
                    st.rerun()

        df_t2_flat = edited_t2_df

        st.divider()
        st.subheader("🖨️ 各種出力形式の選択（PDF ＆ Excel）")
        
        # 出力形式のセレクトボックスを追加
        output_format = st.radio("出力形式を選択してください", ["PDFで出力する", "Excel (.xlsx) で出力する"], horizontal=True, key="out_fmt_select")
        
        col_out1, col_out2 = st.columns(2)
        
        # ==========================================
        # ① 納品管理帳（予定表）の出力（PDF or Excel）
        # ==========================================
        with col_out1:
            if output_format == "PDFで出力する":
                if st.button("🚀 顧客別・便別 納品管理帳(予定表)PDFを生成 [A4横]", type="secondary", key="btn_generate_schedule_pdf"):
                    buffer_sch = io.BytesIO()
                    doc_sch = SimpleDocTemplate(buffer_sch, pagesize=landscape(A4), rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
                    story_sch = []
                    
                    styles = getSampleStyleSheet()
                    title_style = ParagraphStyle('JpTitle', parent=styles['Heading1'], fontName=JP_FONT, fontSize=15, leading=20, alignment=1)
                    heading_style = ParagraphStyle('JpHeading', parent=styles['Heading2'], fontName=JP_FONT, fontSize=11, leading=15)
                    cell_style = ParagraphStyle('JpCell', parent=styles['Normal'], fontName=JP_FONT, fontSize=8, leading=11)
                    cell_right = ParagraphStyle('JpCellR', parent=styles['Normal'], fontName=JP_FONT, fontSize=8, leading=11, alignment=2)

                    story_sch.append(Paragraph(f"<b>納品管理帳（予定表） （予定日: {target_date_str}）</b>", title_style))
                    story_sch.append(Spacer(1, 6*mm))
                    
                    groups = list(df_t2_flat.groupby(["出版社名", "配車区分"]))
                    for idx, ((pub, vehicle), group) in enumerate(groups):
                        story_sch.append(Paragraph(f"<b>■ 宛先（顧客）: {pub}  ／  配車便: {vehicle}</b>", heading_style))
                        story_sch.append(Spacer(1, 2*mm))
                        
                        table_data = [[
                            Paragraph("<b>判型 / 書名 / ISBN</b>", cell_style),
                            Paragraph("<b>刷回</b>", cell_style),
                            Paragraph("<b>定価</b>", cell_style),
                            Paragraph("<b>印刷部数</b>", cell_style),
                            Paragraph("<b>納品部数</b>", cell_style),
                            Paragraph("<b>見本</b>", cell_style),
                            Paragraph("<b>総部数</b>", cell_style),
                            Paragraph("<b>単価</b>", cell_style),
                            Paragraph("<b>預り在庫</b>", cell_style),
                            Paragraph("<b>売上(納品金額)</b>", cell_style)
                        ]]
                        
                        for _, row in group.iterrows():
                            n_qty = safe_float(row.get("納品部数", 0))
                            m_qty = safe_float(row.get("見本", 0))
                            tanka = safe_float(row.get("単価", 0))
                            total_qty = int(n_qty + m_qty)
                            uriage = int(total_qty * tanka)

                            format_val = row.get('判型', '')
                            title_val = row.get('書名', '')
                            isbn_val = row.get('ISBN', '')
                            
                            book_info_html = f"[{format_val}] {title_val}<br/>(ISBN: {isbn_val})"

                            table_data.append([
                                Paragraph(book_info_html, cell_style),
                                Paragraph(str(row.get("刷回", "")), cell_style),
                                Paragraph(str(row.get("定価", "")), cell_right),
                                Paragraph(str(row.get("印刷部数", "")), cell_right),
                                Paragraph(str(row.get("納品部数", "")), cell_right),
                                Paragraph(str(row.get("見本", "")), cell_right),
                                Paragraph(str(total_qty), cell_right),
                                Paragraph(str(row.get("単価", "")), cell_right),
                                Paragraph(str(row.get("預かり在庫", "")), cell_right),
                                Paragraph(f"{uriage:,}", cell_right),
                            ])
                        
                        schedule_table = Table(table_data, colWidths=[87*mm, 15*mm, 18*mm, 18*mm, 18*mm, 14*mm, 18*mm, 18*mm, 20*mm, 30*mm])
                        schedule_table.setStyle(TableStyle([
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('TOPPADDING', (0,0), (-1,-1), 4),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                        ]))
                        
                        story_sch.append(schedule_table)
                        story_sch.append(Spacer(1, 6*mm))
                        
                        is_iwanami = ("岩波" in pub)
                        if is_iwanami and idx < len(groups) - 1:
                            story_sch.append(PageBreak())

                    doc_sch.build(story_sch)
                    buffer_sch.seek(0)
                    
                    st.download_button(
                        label="📥 納品管理帳(予定表)PDFをダウンロード [A4横]",
                        data=buffer_sch,
                        file_name=f"納品管理帳_{target_date_str}.pdf",
                        mime="application/pdf",
                        key="dl_sch_pdf"
                    )
            else:
                # Excel出力の処理
                if st.button("🚀 顧客別・便別 納品管理帳(予定表)Excelを生成 [.xlsx]", type="secondary", key="btn_generate_schedule_xlsx"):
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "納品管理帳"
                    
                    headers = ["出版社名", "配車便", "判型", "書名", "ISBN", "刷回", "定価", "印刷部数", "納品部数", "見本", "総部数", "単価", "預り在庫", "売上(納品金額)", "予定日"]
                    ws.append(headers)
                    
                    for _, row in df_t2_flat.iterrows():
                        n_qty = safe_float(row.get("納品部数", 0))
                        m_qty = safe_float(row.get("見本", 0))
                        tanka = safe_float(row.get("単価", 0))
                        total_qty = int(n_qty + m_qty)
                        uriage = int(total_qty * tanka)
                        
                        ws.append([
                            row.get("出版社名", ""),
                            row.get("配車区分", ""),
                            row.get("判型", ""),
                            row.get("書名", ""),
                            row.get("ISBN", ""),
                            row.get("刷回", ""),
                            safe_float(row.get("定価", 0)),
                            safe_float(row.get("印刷部数", 0)),
                            safe_float(row.get("納品部数", 0)),
                            safe_float(row.get("見本", 0)),
                            total_qty,
                            tanka,
                            safe_float(row.get("預かり在庫", 0)),
                            uriage,
                            target_date_str
                        ])
                    
                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 納品管理帳(予定表)Excelをダウンロード",
                        data=excel_buffer,
                        file_name=f"納品管理帳_{target_date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sch_xlsx"
                    )

        # ==========================================
        # ② 専用伝票の出力（PDF専用）
        # ==========================================
        with col_out2:
            if st.button("🚀 【専用伝票】項目・罫線なし・9ptで生成", type="primary", key="btn_generate_slips"):
                grouped = df_t2_flat.groupby(["出版社名", "配車区分"])
                
                updated_count = 0
                for (pub, vehicle), group in grouped:
                    max_slip_no = 0
                    for r in df_flat.to_dict(orient="records"):
                        if r["出版社名"] == pub and r["配車区分"] == vehicle:
                            s_no = str(r["専用伝票番号"])
                            if s_no.isdigit():
                                max_slip_no = max(max_slip_no, int(s_no))
                    
                    current_no = max(1, max_slip_no + 1)
                    
                    for idx, row in group.iterrows():
                        if not str(row["専用伝票番号"]).isdigit():
                            df_t2_flat.at[idx, "専用伝票番号"] = str(current_no)
                            updated_count += 1
                
                st.success(f"🔒 番号死守ガード適用（新規採番: {updated_count}件）")
                
                buffer_slip = io.BytesIO()
                doc_slip = SimpleDocTemplate(buffer_slip, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
                story_slip = []
                
                styles = getSampleStyleSheet()
                slip_cell_style = ParagraphStyle('SlipCell9', parent=styles['Normal'], fontName=JP_FONT, fontSize=9, leading=12)
                slip_header_style = ParagraphStyle('SlipHeader9', parent=styles['Normal'], fontName=JP_FONT, fontSize=9, leading=12)
                
                for (pub, vehicle), group in df_t2_flat.groupby(["出版社名", "配車区分"]):
                    items = group.to_dict(orient="records")
                    slip_no = items[0]["専用伝票番号"] if items else "1"
                    
                    CHUNK_SIZE = 6
                    for i in range(0, len(items), CHUNK_SIZE):
                        chunk = items[i:i + CHUNK_SIZE]
                        
                        header_text = f"宛先: {pub}  便: {vehicle}  伝票No: {slip_no}  予定日: {target_date_str}"
                        story_slip.append(Paragraph(header_text, slip_header_style))
                        story_slip.append(Spacer(1, 15*mm))
                        
                        page_rows = []
                        for item in chunk:
                            row1 = [
                                Paragraph(str(item.get("受注番号", "")), slip_cell_style),
                                Paragraph(f"[{item.get('判型', '')}] {item.get('書名', '')} (ISBN: {item.get('ISBN', '')})", slip_cell_style),
                                Paragraph(f"定:{item.get('定価', '')} 刷:{item.get('刷回', '')}", slip_cell_style),
                                Paragraph(str(item.get("単価", "")), slip_cell_style),
                                Paragraph(str(item.get("納品部数", "")), slip_cell_style),
                                Paragraph(str(item.get("PL数", "")), slip_cell_style),
                                Paragraph(str(item.get("見本", "")), slip_cell_style)
                            ]
                            row2 = [
                                Paragraph(str(item.get("摘要", "")), slip_cell_style),
                                Paragraph(f"預:{item.get('預かり在庫', '')}", slip_cell_style),
                                Paragraph("✓" if item.get("在庫納品") else "", slip_cell_style),
                                "", "", "", ""
                            ]
                            page_rows.append(row1)
                            page_rows.append(row2)
                            
                        while len(page_rows) < 14:
                            page_rows.append(["", "", "", "", "", "", ""])
                            
                        slip_table = Table(page_rows, colWidths=[25*mm, 65*mm, 20*mm, 15*mm, 15*mm, 12*mm, 13*mm])
                        slip_table.setStyle(TableStyle([
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('TOPPADDING', (0,0), (-1,-1), 2),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
                        ]))
                        
                        story_slip.append(slip_table)
                        story_slip.append(PageBreak())
                
                doc_slip.build(story_slip)
                buffer_slip.seek(0)
                
                st.download_button(
                    label="📥 専用伝票PDFをダウンロード [項目・罫線なし/9pt]",
                    data=buffer_slip,
                    file_name=f"納品専用伝票_{target_date_str}.pdf",
                    mime="application/pdf",
                    key="dl_slip_pdf"
                )
    else:
        st.info(f"ℹ️ 指定された予定日【 {target_date_str} 】（field-90_8）に該当するデータはありません。")
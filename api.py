import io
import os
import json
import requests
import pandas as pd
from datetime import datetime, timedelta, date
import streamlit as st
import copy

# ReportLab関連（PDF出力用）
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
                # レコードIDの正確な取得
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
        h_isbn = str(r.get("field-93", ""))
        h_qty = str(r.get("field-4", ""))
        h_price = str(r.get("field-20", ""))
        h_print_qty = str(r.get("field-3", ""))
        h_deposit = str(r.get("field-81", ""))
        
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
            
            item_89 = sub_89[i] if isinstance(sub_89, list) and len(sub_89) > i else {}
            sub_id_89 = item_89.get("recordSubId", item_89.get("Id", ""))
            rec_89 = item_89.get("record", item_89) if isinstance(item_89, dict) else {}
            
            flat.append({
                "更新": False,
                "_record_idx": record_idx,
                "_sub_idx": i,
                "_id_90": sub_id_90,
                "_id_88": sub_id_88,
                "_id_89": sub_id_89,
                "受注番号": h_ord_no,
                "受注日": h_ord_date,
                "出版社名": h_pub,
                "書名": h_title,
                "ISBN": h_isbn,
                "製本数": h_qty,
                "定価": h_price,
                "印字数": h_print_qty,
                "預かり数": h_deposit,
                "予定日": parse_date_safe(rec_90.get("field-90_8")),
                "納品日": parse_date_safe(rec_88.get("field-88_2")),
                "納品数": str(rec_88.get("field-88_4", "")),
                "摘要": str(rec_88.get("field-88_5", "")),
                "在庫納品": bool(rec_88.get("field-88_6", False)),
                "PL数": str(rec_88.get("field-88_7", "")),
                "見本数": str(rec_88.get("field-88_8", "")),
                "専用伝票番号": str(rec_88.get("field-88_9", "")),
                "配車区分": str(rec_88.get("field-88_10", "") if rec_88.get("field-88_10") else "便なし"),
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
                "_record_idx": None, "_sub_idx": None, "_id_90": None, "_id_88": None, "_id_89": None,
                "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
                "受注日": st.column_config.DateColumn("受注日", format="YYYY-MM-DD", disabled=True),
                "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
                "書名": st.column_config.TextColumn("書名", disabled=True),
                "ISBN": st.column_config.TextColumn("ISBN", disabled=True),
                "製本数": st.column_config.TextColumn("製本数", disabled=True),
                "定価": st.column_config.TextColumn("定価", disabled=True),
                "印字数": st.column_config.TextColumn("印字数", disabled=True),
                "預かり数": st.column_config.TextColumn("預かり数", disabled=True),
                "予定日": st.column_config.DateColumn("予定日(90_8)", format="YYYY-MM-DD", disabled=True),
                "納品日": st.column_config.DateColumn("納品日", format="YYYY-MM-DD"),
                "納品数": st.column_config.TextColumn("納品数"),
                "摘要": st.column_config.TextColumn("摘要"),
                "在庫納品": st.column_config.CheckboxColumn("在庫納品"),
                "PL数": st.column_config.TextColumn("PL数"),
                "見本数": st.column_config.TextColumn("見本数"),
                "専用伝票番号": st.column_config.TextColumn("伝票番号"),
                "配車区分": st.column_config.SelectboxColumn("配車区分", options=["", "便なし", "A車", "B車", "C車", "チャーター"]),
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
                        st.error(f"❌ レコードIDが見つかりません (index: {r_idx})")
                        continue
                    
                    update_url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records/{record_id}"
                    first_row = group_df.iloc[0]
                    record_payload = {
                        "field-2": str(first_row["受注番号"]),
                        "field-1": str(first_row["受注日"]) if pd.notna(first_row["受注日"]) else "",
                        "field-4": str(first_row["製本数"]),
                        "field-20": str(first_row["定価"]),
                        "field-3": str(first_row["印字数"]),
                        "field-81": str(first_row["預かり数"]),
                        "field-93": str(first_row["ISBN"]),
                    }
                    
                    orig_sub_88 = original_rec.get("field-88", [])
                    sub88_list = []
                    for _, row in group_df.iterrows():
                        sub_i = int(row["_sub_idx"])
                        item_88 = copy.deepcopy(orig_sub_88[sub_i]) if (isinstance(orig_sub_88, list) and len(orig_sub_88) > sub_i) else {}
                        rec_dict = item_88.get("record", {}) if isinstance(item_88, dict) else {}
                        
                        rec_dict["field-88_2"] = str(row["納品日"]) if pd.notna(row["納品日"]) else ""
                        rec_dict["field-88_4"] = str(row["納品数"])
                        rec_dict["field-88_5"] = str(row["摘要"])
                        rec_dict["field-88_6"] = bool(row["在庫納品"])
                        rec_dict["field-88_7"] = str(row["PL数"])
                        rec_dict["field-88_8"] = str(row["見本数"])
                        rec_dict["field-88_9"] = str(row["専用伝票番号"])
                        rec_dict["field-88_10"] = str(row["配車区分"])
                        
                        item_88["record"] = rec_dict
                        sub88_list.append(item_88)
                    
                    record_payload["field-88"] = sub88_list
                    payload = {"apiVersion": 1.1, "record": record_payload}
                    
                    try:
                        res = requests.put(update_url, headers=headers, json=payload, timeout=15)
                        if res.status_code in [200, 204]:
                            success_count += 1
                        else:
                            # ★ 失敗した原因を画面に赤字で詳しく表示するデバッグ
                            st.error(f"❌ 書き戻しエラー (ID: {record_id}, ステータス: {res.status_code}): {res.text}")
                    except Exception as ex:
                        st.error(f"🚨 例外発生: {ex}")
                
                if success_count > 0:
                    st.success(f"✅ {success_count} 件のレコードを正常に上書き保存しました！")
                    st.session_state["api_records"] = fetch_pocket_records(start_date=t1_start, end_date=t1_end)
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
                "_record_idx": None, "_sub_idx": None, "_id_90": None, "_id_88": None, "_id_89": None,
                "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
                "受注日": st.column_config.DateColumn("受注日", format="YYYY-MM-DD", disabled=True),
                "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
                "書名": st.column_config.TextColumn("書名", disabled=True),
                "ISBN": st.column_config.TextColumn("ISBN", disabled=True),
                "製本数": st.column_config.TextColumn("製本数", disabled=True),
                "定価": st.column_config.TextColumn("定価", disabled=True),
                "印字数": st.column_config.TextColumn("印字数", disabled=True),
                "預かり数": st.column_config.TextColumn("預かり数", disabled=True),
                "予定日": st.column_config.DateColumn("予定日", format="YYYY-MM-DD", disabled=True),
                "納品日": st.column_config.DateColumn("納品日", format="YYYY-MM-DD"),
                "納品数": st.column_config.TextColumn("納品数"),
                "PL数": st.column_config.TextColumn("PL数"),
                "見本数": st.column_config.TextColumn("見本数"),
                "配車区分": st.column_config.SelectboxColumn("配車区分", options=["", "便なし", "A車", "B車", "C車", "チャーター"]),
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
                        st.error(f"❌ レコードIDが見つかりません (index: {r_idx})")
                        continue
                    
                    update_url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records/{record_id}"
                    first_row = group_df.iloc[0]
                    record_payload = {
                        "field-2": str(first_row["受注番号"]),
                        "field-1": str(first_row["受注日"]) if pd.notna(first_row["受注日"]) else "",
                        "field-4": str(first_row["製本数"]),
                        "field-20": str(first_row["定価"]),
                        "field-3": str(first_row["印字数"]),
                        "field-81": str(first_row["預かり数"]),
                        "field-93": str(first_row["ISBN"]),
                    }
                    
                    orig_sub_88 = original_rec.get("field-88", [])
                    sub88_list = []
                    for _, row in group_df.iterrows():
                        sub_i = int(row["_sub_idx"])
                        item_88 = copy.deepcopy(orig_sub_88[sub_i]) if (isinstance(orig_sub_88, list) and len(orig_sub_88) > sub_i) else {}
                        rec_dict = item_88.get("record", {}) if isinstance(item_88, dict) else {}
                        
                        rec_dict["field-88_2"] = str(row["納品日"]) if pd.notna(row["納品日"]) else ""
                        rec_dict["field-88_4"] = str(row["納品数"])
                        rec_dict["field-88_5"] = str(row["摘要"])
                        rec_dict["field-88_6"] = bool(row["在庫納品"])
                        rec_dict["field-88_7"] = str(row["PL数"])
                        rec_dict["field-88_8"] = str(row["見本数"])
                        rec_dict["field-88_9"] = str(row["専用伝票番号"])
                        rec_dict["field-88_10"] = str(row["配車区分"])
                        
                        item_88["record"] = rec_dict
                        sub88_list.append(item_88)
                    
                    record_payload["field-88"] = sub88_list
                    payload = {"apiVersion": 1.1, "record": record_payload}
                    
                    try:
                        res = requests.put(update_url, headers=headers, json=payload, timeout=15)
                        if res.status_code in [200, 204]:
                            success_count += 1
                        else:
                            # ★ 失敗した原因を画面に赤字で詳しく表示するデバッグ
                            st.error(f"❌ タブ2書き戻しエラー (ID: {record_id}, ステータス: {res.status_code}): {res.text}")
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
        st.subheader("🖨️ 各種出力（納品予定表 ＆ 専用伝票）")
        
        col_out1, col_out2 = st.columns(2)
        
        # --- ① 納品予定表PDFの生成 ---
        with col_out1:
            if st.button("🚀 顧客別・便別 納品予定表PDFを生成", type="secondary", key="btn_generate_schedule_pdf"):
                buffer_sch = io.BytesIO()
                doc_sch = SimpleDocTemplate(buffer_sch, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
                story_sch = []
                
                styles = getSampleStyleSheet()
                
                font_registered = False
                for font_path in ["NotoSansJP-Regular.ttf", "fonts/NotoSansJP-Regular.ttf", "/usr/share/fonts/opentype/noto/NotoSansJP-Regular.ttf"]:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont("NotoSansJP", font_path))
                            font_registered = True
                            break
                        except Exception:
                            pass
                
                jp_font_name = "NotoSansJP" if font_registered else "Helvetica"
                
                title_style = ParagraphStyle('JpTitle', parent=styles['Heading1'], fontName=jp_font_name, fontSize=16, leading=22, alignment=1)
                heading_style = ParagraphStyle('JpHeading', parent=styles['Heading2'], fontName=jp_font_name, fontSize=11, leading=16)
                cell_style = ParagraphStyle('JpCell', parent=styles['Normal'], fontName=jp_font_name, fontSize=9, leading=12)

                story_sch.append(Paragraph(f"<b>納品予定表 （予定日: {target_date_str}）</b>", title_style))
                story_sch.append(Spacer(1, 8*mm))
                
                for (pub, vehicle), group in df_t2_flat.groupby(["出版社名", "配車区分"]):
                    story_sch.append(Paragraph(f"<b>■ 宛先（顧客）: {pub}  ／  配車便: {vehicle}</b>", heading_style))
                    story_sch.append(Spacer(1, 2*mm))
                    
                    table_data = [[
                        Paragraph("<b>受注番号</b>", cell_style),
                        Paragraph("<b>書名</b>", cell_style),
                        Paragraph("<b>納品数</b>", cell_style),
                        Paragraph("<b>PL数</b>", cell_style),
                        Paragraph("<b>見本数</b>", cell_style),
                        Paragraph("<b>摘要</b>", cell_style)
                    ]]
                    
                    for _, row in group.iterrows():
                        table_data.append([
                            Paragraph(str(row.get("受注番号", "")), cell_style),
                            Paragraph(str(row.get("書名", "")), cell_style),
                            Paragraph(str(row.get("納品数", "")), cell_style),
                            Paragraph(str(row.get("PL数", "")), cell_style),
                            Paragraph(str(row.get("見本数", "")), cell_style),
                            Paragraph(str(row.get("摘要", "")), cell_style),
                        ])
                    
                    schedule_table = Table(table_data, colWidths=[25*mm, 75*mm, 20*mm, 15*mm, 15*mm, 30*mm])
                    schedule_table.setStyle(TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('TOPPADDING', (0,0), (-1,-1), 4),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                    ]))
                    
                    story_sch.append(schedule_table)
                    story_sch.append(Spacer(1, 6*mm))
                
                doc_sch.build(story_sch)
                buffer_sch.seek(0)
                
                st.download_button(
                    label="📥 納品予定表PDFをダウンロード",
                    data=buffer_sch,
                    file_name=f"納品予定表_{target_date_str}.pdf",
                    mime="application/pdf",
                    key="dl_sch_pdf"
                )

        # --- ② 専用伝票PDFの生成（番号死守ガード付き） ---
        with col_out2:
            if st.button("🚀 【番号死守】専用伝票PDFを生成", type="primary", key="btn_generate_slips"):
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
                normal_style = styles['Normal']
                
                for (pub, vehicle), group in df_t2_flat.groupby(["出版社名", "配車区分"]):
                    items = group.to_dict(orient="records")
                    slip_no = items[0]["専用伝票番号"] if items else "1"
                    
                    CHUNK_SIZE = 6
                    for i in range(0, len(items), CHUNK_SIZE):
                        chunk = items[i:i + CHUNK_SIZE]
                        
                        header_text = f"<b>【専用伝票】 宛先: {pub} 御中  配車便: {vehicle}  伝票番号: 第 {slip_no} 号  予定日: {target_date_str}</b>"
                        story_slip.append(Paragraph(header_text, normal_style))
                        story_slip.append(Spacer(1, 5*mm))
                        
                        page_rows = [
                            ["受注番号", "書名 / ISBN", "定価", "刷数", "納品数", "PL数", "見本数"],
                            ["備考 / 摘要", "預かり数 / 印字数", "在庫納品", "", "", "", ""]
                        ]
                        
                        for item in chunk:
                            row1 = [
                                str(item.get("受注番号", "")),
                                f"{item.get('書名', '')}\n(ISBN: {item.get('ISBN', '')})",
                                str(item.get("定価", "")),
                                str(item.get("製本数", "")),
                                str(item.get("納品数", "")),
                                str(item.get("PL数", "")),
                                str(item.get("見本数", ""))
                            ]
                            row2 = [
                                str(item.get("摘要", "")),
                                f"預:{item.get('預かり数', '')} / 印:{item.get('印字数', '')}",
                                "✓" if item.get("在庫納品") else "",
                                "", "", "", ""
                            ]
                            page_rows.append(row1)
                            page_rows.append(row2)
                            
                        while len(page_rows) < 14:
                            page_rows.append(["", "", "", "", "", "", ""])
                            
                        slip_table = Table(page_rows, colWidths=[25*mm, 65*mm, 15*mm, 15*mm, 15*mm, 15*mm, 15*mm])
                        slip_table.setStyle(TableStyle([
                            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                            ('FONTSIZE', (0,0), (-1,-1), 9),
                        ]))
                        
                        story_slip.append(slip_table)
                        story_slip.append(PageBreak())
                
                doc_slip.build(story_slip)
                buffer_slip.seek(0)
                
                st.download_button(
                    label="📥 専用伝票PDFをダウンロード",
                    data=buffer_slip,
                    file_name=f"納品専用伝票_{target_date_str}.pdf",
                    mime="application/pdf",
                    key="dl_slip_pdf"
                )
    else:
        st.info(f"ℹ️ 指定された予定日【 {target_date_str} 】（field-90_8）に該当するデータはありません。")
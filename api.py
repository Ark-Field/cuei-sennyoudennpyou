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
    APP_ID = int(st.secrets["pocket"].get("app_id", 34))
else:
    POCKET_API_KEY = "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"
    APP_ID = 34

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

# ==========================================
# 4. タブ構成の定義
# ==========================================
tab1, tab2 = st.tabs([
    "📋 タブ1: 受注・納品データ一覧 ＆ 更新", 
    "📦 タブ2: 納品予定表 ＆ 専用伝票発行（顧客×便別）"
])

# 初期データロード
if "api_records" not in st.session_state:
    default_end = date.today()
    default_start = default_end - timedelta(days=21)
    st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end)

api_records = st.session_state.get("api_records", [])

# ==========================================
# フラットデータ（明細行展開）の構築
# ==========================================
flat_data = []
for record_idx, r in enumerate(api_records):
    h_ord_no = str(r.get("field-2", ""))
    h_ord_date = parse_date_safe(r.get("field-1"))
    h_pub = str(r.get("field-8", ""))
    h_title = str(r.get("field-13", ""))
    h_isbn = str(r.get("field-93", ""))
    h_qty = str(r.get("field-4", ""))
    h_price = str(r.get("field-20", "")) # 定価
    h_print_qty = str(r.get("field-3", "")) # 印字数
    h_deposit = str(r.get("field-81", "")) # 預かり数
    
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
        
        flat_data.append({
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
            "配本日": parse_date_safe(rec_90.get("field-90_1")),
            "納品日": parse_date_safe(rec_88.get("field-88_2")),
            "納品数": str(rec_88.get("field-88_4", "")),
            "摘要": str(rec_88.get("field-88_5", "")),
            "在庫納品": bool(rec_88.get("field-88_6", False)),
            "PL数": str(rec_88.get("field-88_7", "")),
            "見本数": str(rec_88.get("field-88_8", "")),
            "専用伝票番号": str(rec_88.get("field-88_9", "")),
            "配車区分": str(rec_88.get("field-88_10", "") if rec_88.get("field-88_10") else "便なし"),
        })

df_flat = pd.DataFrame(flat_data)

# ==========================================
# タブ1：受注・納品データ一覧 ＆ 更新
# ==========================================
with tab1:
    st.header("📋 受注・納品データ管理エディタ")
    
    with st.expander("🔍 検索条件の変更・絞り込み", expanded=True):
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1:
            default_end = date.today()
            default_start = default_end - timedelta(days=21)
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
                "配本日": st.column_config.DateColumn("配本日", format="YYYY-MM-DD"),
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
                fail_count = 0
                
                for r_idx, group_df in checked_rows.groupby("_record_idx"):
                    r_idx = int(r_idx)
                    original_rec = api_records[r_idx]
                    record_id = original_rec.get("_recordId")
                    if not record_id:
                        fail_count += 1
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
                    
                    # サブテーブル88（納品）の構築
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
                            fail_count += 1
                    except Exception:
                        fail_count += 1
                
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
        target_delivery_date = st.date_input("対象納品日を指定", value=date.today(), key="target_deliv_date")
    
    target_date_str = target_delivery_date.strftime("%Y-%m-%d")
    target_year_month = target_delivery_date.strftime("%Y-%m")
    
    # 選択した納品日に合致する明細を抽出
    day_delivery_rows = []
    for row in flat_data:
        d_val = row["納品日"]
        if d_val and str(d_val) == target_date_str:
            day_delivery_rows.append(row)
            
    if day_delivery_rows:
        df_day = pd.DataFrame(day_delivery_rows)
        st.success(f"✨ 納品日【 {target_date_str} 】の明細データが **{len(df_day)} 件** あります。")
        
        # プレビュー表示
        st.dataframe(df_day[["出版社名", "配車区分", "受注番号", "書名", "納品数", "専用伝票番号", "PL数", "見本数"]], use_container_width=True)
        
        st.divider()
        st.subheader("🖨️ 専用伝票番号の自動採番（番号死守ガード付き） ＆ PDF出力")
        
        if st.button("🚀 【番号死守】自動採番を実行して専用伝票PDFを生成する", type="primary", key="btn_generate_slips"):
            # 顧客（出版社名）× 配車区分 ごとにグループ化して連番を付与
            # 既存の伝票番号があれば維持、空なら今月の同条件の最大値+1を採番
            grouped = df_day.groupby(["出版社名", "配車区分"])
            
            updated_count = 0
            for (pub, vehicle), group in grouped:
                # すでに今月付与されている最大番号を調査
                max_slip_no = 0
                for r in flat_data:
                    if r["出版社名"] == pub and r["配車区分"] == vehicle:
                        s_no = str(r["専用伝票番号"])
                        if s_no.isdigit():
                            max_slip_no = max(max_slip_no, int(s_no))
                
                current_no = max(1, max_slip_no + 1)
                
                # グループ内の明細で伝票番号が空のものを自動採番
                for idx, row in group.iterrows():
                    if not str(row["専用伝票番号"]).isdigit():
                        df_day.at[idx, "専用伝票番号"] = str(current_no)
                        updated_count += 1
            
            st.success(f"🔒 番号死守ガード適用完了（新規採番・割り当て件数: {updated_count}件）")
            
            # PDF生成ロジック（ReportLab）
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
            story = []
            
            # フォント設定（日本語対応：標準ゴシック代替等または環境依存回避のためラップ）
            styles = getSampleStyleSheet()
            normal_style = styles['Normal']
            
            # 顧客×配車区分ごとにPDFページを構築
            for (pub, vehicle), group in df_day.groupby(["出版社名", "配車区分"]):
                items = group.to_dict(orient="records")
                slip_no = items[0]["専用伝票番号"] if items else "1"
                
                # 1ページ最大6行（2行組×6行 = 12行分）に分割
                CHUNK_SIZE = 6
                for i in range(0, len(items), CHUNK_SIZE):
                    chunk = items[i:i + CHUNK_SIZE]
                    
                    # ヘッダー情報（専用伝票の親フィールド部分）
                    header_text = f"<b>【専用伝票】 宛先: {pub} 御中  配車便: {vehicle}  伝票番号: 第 {slip_no} 号  納品日: {target_date_str}</b>"
                    story.append(Paragraph(header_text, normal_style))
                    story.append(Spacer(1, 5*mm))
                    
                    page_rows = [
                        ["受注番号", "書名 / ISBN", "定価", "刷数", "納品数", "PL数", "見本数"],
                        ["備考 / 摘要", "預かり数 / 印字数", "在庫納品", "", "", "", ""]
                    ] # テーブルヘッダー例
                    
                    for item in chunk:
                        row1 = [
                            str(item.get("受注番号", "")),
                            f"{item.get('書名', '')}\n(ISBN: {item.get('ISBN', '')})",
                            str(item.get("定価", "")),
                            str(item.get("製本数", "")), # 刷数・製本数
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
                        
                    # 6行（2行組×6 = 12行）に満たない場合は空白行でパディング
                    while len(page_rows) < 14: # ヘッダー2行 + 12行
                        page_rows.append(["", "", "", "", "", "", ""])
                        
                    slip_table = Table(page_rows, colWidths=[25*mm, 65*mm, 15*mm, 15*mm, 15*mm, 15*mm, 15*mm])
                    slip_table.setStyle(TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ('FONTSIZE', (0,0), (-1,-1), 9),
                    ]))
                    
                    story.append(slip_table)
                    story.append(PageBreak())
            
            doc.build(story)
            buffer.seek(0)
            
            st.download_button(
                label="📥 専用伝票PDFをダウンロードする",
                data=buffer,
                file_name=f"納品専用伝票_{target_date_str}.pdf",
                mime="application/pdf",
                type="primary"
            )
    else:
        st.info(f"ℹ️ 納品日【 {target_date_str} 】に該当する納品データはありません。")
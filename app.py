"""Kamu Satın Alma Yönetim Sistemi — Dosya Bazlı Minimalist Uygulama."""

from __future__ import annotations

import datetime as dt
import html as html_lib
import importlib
import json
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

_APP_DIR = str(Path(__file__).resolve().parent)
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

import ai_parser as _ai_parser_mod
import db as _db_mod
import doc_generator as _doc_gen_mod
import search_engine as _search_mod

if not hasattr(_ai_parser_mod, "verify_suppliers_with_ai"):
    _ai_parser_mod = importlib.reload(_ai_parser_mod)
if not hasattr(_db_mod, "increment_supplier_trust"):
    _db_mod = importlib.reload(_db_mod)
if not hasattr(_doc_gen_mod, "build_market_research_docx"):
    _doc_gen_mod = importlib.reload(_doc_gen_mod)
if not hasattr(_search_mod, "run_three_ring_search"):
    _search_mod = importlib.reload(_search_mod)

from ai_parser import (
    parse_demand_form,
    parse_mkys_excel,
    parse_offer_document,
    parse_tr_number,
    sanitize_pool_batch_with_ai,
)
from db import (
    add_case_suppliers,
    apply_pool_sanitization_results,
    create_case,
    delete_case,
    delete_case_offer,
    delete_case_suppliers,
    evaluate_case_offers,
    get_case,
    get_case_items,
    get_case_offers,
    get_case_suppliers,
    get_mkys_stats,
    get_unverified_pool_records,
    import_mkys_records,
    increment_supplier_trust,
    init_db,
    list_cases,
    save_default_commission,
    save_item_intel,
    seed_unverified_from_xlsx,
    sync_case_items,
    update_case_meta,
    update_supplier_selections_and_edits,
    upsert_case_offer,
)
from doc_generator import (
    build_invitation_docx,
    build_invitation_text,
    build_market_research_docx,
    build_rfq_draft,
    format_qty,
    format_tr_number,
)
from search_engine import (
    CACHE_XLSX,
    DEFAULT_MAX_SITES,
    analyze_item_with_gemini,
    build_excel_bytes,
    collect_bcc_emails,
    dataframe_to_utf8_bom_csv,
    get_secret,
    infer_product_groups,
    load_ai_cache,
    load_cache,
    run_three_ring_search,
    save_cache,
)


# ---------------------------------------------------------------------------
# Minimalist Açık Tema CSS
# ---------------------------------------------------------------------------

def inject_minimal_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            color: #0F172A;
        }
        .stApp {
            background-color: #F8FAFC;
            color: #0F172A;
        }
        section[data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 1px solid #E2E8F0;
        }
        section[data-testid="stSidebar"] .block-container {
            padding-top: 1.25rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }

        /* Gereksiz Streamlit chrome alanlarını gizle (sidebar açma düğmesini koru) */
        #MainMenu, footer, [data-testid="stToolbar"], .stDeployButton,
        [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
            display: none !important;
            visibility: hidden !important;
        }
        [data-testid="stHeader"] {
            background: transparent !important;
            height: 2.2rem !important;
        }

        .block-container {
            padding-top: 1.4rem !important;
            padding-bottom: 2.5rem !important;
            padding-left: 2.2rem !important;
            padding-right: 2.2rem !important;
            max-width: 1320px !important;
        }

        /* Üst Başlık ve Kartlar */
        .app-header {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-left: 4px solid #0F766E;
            border-radius: 12px;
            padding: 18px 24px;
            margin-bottom: 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 16px;
        }
        .app-header h1 {
            margin: 0;
            font-size: 1.32rem;
            font-weight: 700;
            color: #0F172A;
            letter-spacing: -0.01em;
        }
        .app-header p {
            margin: 4px 0 0 0;
            font-size: 0.88rem;
            color: #64748B;
        }
        .type-badge {
            display: inline-block;
            background: #F0FDFA;
            color: #0F766E;
            border: 1px solid #CCFBF1;
            border-radius: 999px;
            padding: 5px 14px;
            font-size: 0.8rem;
            font-weight: 600;
            white-space: nowrap;
        }

        .clean-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 18px 20px;
            margin-bottom: 16px;
        }
        .card-title {
            font-size: 0.98rem;
            font-weight: 600;
            color: #0F172A;
            margin: 0 0 6px 0;
        }
        .card-sub {
            font-size: 0.84rem;
            color: #64748B;
            margin: 0 0 12px 0;
        }

        .kpi-box {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-top: 3px solid #0F766E;
            border-radius: 10px;
            padding: 12px 16px;
            margin-bottom: 14px;
        }
        .kpi-lbl {
            font-size: 0.74rem;
            font-weight: 600;
            text-transform: uppercase;
            color: #64748B;
            letter-spacing: 0.04em;
        }
        .kpi-val {
            font-size: 1.38rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 2px;
        }

        .pill {
            display: inline-block;
            background: #F0FDFA;
            color: #0F766E;
            border: 1px solid #CCFBF1;
            border-radius: 999px;
            padding: 2px 10px;
            font-size: 0.76rem;
            font-weight: 500;
            margin: 2px 4px 2px 0;
        }
        .letter-preview {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 18px 22px;
            font-family: 'Times New Roman', serif;
            font-size: 1rem;
            line-height: 1.6;
            color: #0F172A;
            text-align: justify;
            margin: 10px 0 16px 0;
        }

        /* Buton Stilleri */
        div.stButton > button[kind="primary"],
        div.stDownloadButton > button {
            background-color: #0F766E !important;
            color: #FFFFFF !important;
            border: 1px solid #0F766E !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
            padding: 0.48rem 1rem !important;
            transition: all 0.15s ease;
        }
        div.stButton > button[kind="primary"]:hover,
        div.stDownloadButton > button:hover {
            background-color: #115E59 !important;
            border-color: #115E59 !important;
            color: #FFFFFF !important;
        }
        div.stButton > button[kind="secondary"] {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            font-size: 0.86rem !important;
            padding: 0.46rem 0.9rem !important;
            transition: all 0.15s ease;
        }
        div.stButton > button[kind="secondary"]:hover {
            border-color: #0F766E !important;
            color: #0F766E !important;
            background-color: #F0FDFA !important;
        }

        /* Sekmeler */
        [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 6px;
            margin-bottom: 14px;
        }
        button[data-baseweb="tab"] {
            border-radius: 8px !important;
            padding: 8px 16px !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
            color: #475569 !important;
            border: none !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #0F766E !important;
            color: #FFFFFF !important;
        }
        [data-baseweb="tab-highlight"] {
            display: none !important;
        }

        /* Tablo ve Girdi Çerçeveleri */
        [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            overflow: hidden;
            background: #FFFFFF;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def check_optional_login() -> bool:
    """APP_PASSWORD .env veya secrets içinde tanımlıysa giriş ister; yoksa doğrudan açılır."""
    expected_pw = get_secret("APP_PASSWORD", "")
    if not expected_pw or st.session_state.get("authenticated"):
        return True
    _, mid, _ = st.columns([1.2, 1.0, 1.2])
    with mid:
        st.markdown(
            '<div class="app-header"><div><h1>🏥 Satın Alma Yönetim Sistemi</h1>'
            '<p>Yetkili Erişimi</p></div></div>',
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            pw = st.text_input("Erişim Şifresi", type="password")
            if st.form_submit_button("Giriş Yap", type="primary", width="stretch"):
                if pw == expected_pw:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Hatalı şifre.")
    return False


def render_copy_button(text_to_copy: str, label: str = "📋 Seçili E-postaları Tek Tıkla Kopyala") -> None:
    """Tarayıcı panosuna tek tıkla kopyalama yapan özel buton."""
    payload_js = json.dumps(text_to_copy)
    label_html = html_lib.escape(label)
    html_code = f"""
    <div style="font-family: Inter, sans-serif; display: flex; align-items: center; gap: 10px;">
      <button id="copyBtn" onclick="copyEmails()" style="
        background-color: #0F766E; color: #FFFFFF; border: none; border-radius: 8px;
        padding: 9px 16px; font-size: 13.5px; font-weight: 600; cursor: pointer;
        font-family: Inter, -apple-system, sans-serif; transition: background 0.15s;
      ">{label_html}</button>
      <span id="copyStatus" style="font-size: 13px; color: #0F766E; font-weight: 600;"></span>
    </div>
    <script>
      function copyEmails() {{
        const txt = {payload_js};
        const btn = document.getElementById('copyBtn');
        const st = document.getElementById('copyStatus');
        if (navigator.clipboard && navigator.clipboard.writeText) {{
          navigator.clipboard.writeText(txt).then(() => {{
            st.textContent = '✓ Panoya kopyalandı!';
            btn.style.backgroundColor = '#115E59';
            setTimeout(() => {{ st.textContent = ''; btn.style.backgroundColor = '#0F766E'; }}, 2200);
          }}).catch(() => fallbackCopy(txt));
        }} else {{
          fallbackCopy(txt);
        }}
      }}
      function fallbackCopy(txt) {{
        const ta = document.createElement('textarea');
        ta.value = txt;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        document.getElementById('copyStatus').textContent = '✓ Panoya kopyalandı!';
      }}
    </script>
    """
    components.html(html_code, height=46)


# ---------------------------------------------------------------------------
# Sol Menü: Alım Dosyaları Yönetimi
# ---------------------------------------------------------------------------

def render_sidebar() -> int:
    with st.sidebar:
        st.markdown("### 📁 Alım Dosyaları")
        st.caption("Her dosyanın kalemleri, firmaları, teklifleri ve tutanağı ayrı saklanır.")

        with st.expander("➕ Yeni Alım Dosyası Oluştur", expanded=False):
            new_title = st.text_input("Dosya Adı", placeholder="Örn: Kardiyoloji Sarf Alımı", key="new_case_title")
            new_type = st.radio("Alım Türü", ["Malzeme", "Hizmet"], horizontal=True, key="new_case_type")
            if st.button("Dosyayı Oluştur", type="primary", width="stretch", key="btn_create_case"):
                cid = create_case(new_title or "Yeni Alım Dosyası", new_type)
                st.session_state["active_case_id"] = cid
                st.toast("Yeni alım dosyası oluşturuldu.")
                st.rerun()

        cases = list_cases()
        if not cases:
            cid = create_case("İlk Alım Dosyası", "Malzeme")
            st.session_state["active_case_id"] = cid
            cases = list_cases()

        active_id = st.session_state.get("active_case_id")
        valid_ids = [int(c["id"]) for c in cases]
        if active_id not in valid_ids:
            active_id = valid_ids[0]
            st.session_state["active_case_id"] = active_id

        st.markdown("---")
        for c in cases:
            cid = int(c["id"])
            is_active = cid == active_id
            icon = "🟢" if is_active else "📄"
            short_title = c["title"] if len(c["title"]) <= 28 else c["title"][:26] + "…"
            meta_txt = f"{c['procurement_type']} · {c['item_count']} kalem · {c['supplier_count']} firma · {c['offer_count']} teklif"
            if st.button(
                f"{icon} {short_title}",
                key=f"case_btn_{cid}",
                type="primary" if is_active else "secondary",
                width="stretch",
                help=f"{c['title']} ({meta_txt})",
            ):
                st.session_state["active_case_id"] = cid
                st.rerun()
            if is_active:
                st.caption(f"↳ {meta_txt}")

        st.markdown("---")
        active_case = next((c for c in cases if int(c["id"]) == active_id), cases[0])
        with st.expander("⚙️ Aktif Dosya İşlemleri", expanded=False):
            confirm_del = st.checkbox("Silme işlemini onayla", key=f"confirm_del_{active_id}")
            if st.button("🗑️ Aktif Dosyayı Sil", disabled=not confirm_del, width="stretch", key=f"del_case_{active_id}"):
                delete_case(active_id)
                st.session_state.pop("active_case_id", None)
                st.toast(f"'{active_case['title']}' silindi.")
                st.rerun()

    return int(st.session_state["active_case_id"])


# ---------------------------------------------------------------------------
# Adım 1: İstem Formu Yükleme ve Kalem Listesi
# ---------------------------------------------------------------------------

def render_step1(case_data: dict[str, Any], items: list[dict[str, Any]], gemini_key: str) -> None:
    case_id = int(case_data["id"])

    col_up, col_meta = st.columns([1.15, 1.0], gap="medium")
    with col_up:
        st.markdown(
            '<div class="clean-card"><div class="card-title">1. Satın Alma İstem Formunu Yükle</div>'
            '<div class="card-sub">PDF, Word (.docx), Excel (.xlsx/.csv) veya görsel (.png/.jpg) yükleyin; Gemini alım adını, türünü ve kalemleri otomatik çıkarsın.</div>',
            unsafe_allow_html=True,
        )
        uploaded_form = st.file_uploader(
            "İstem Formu Seç",
            type=["pdf", "docx", "xlsx", "xlsm", "csv", "png", "jpg", "jpeg", "webp", "txt"],
            key=f"form_uploader_{case_id}",
            label_visibility="collapsed",
        )
        if st.button("✨ Formu Yapay Zeka ile Oku ve Tabloya Aktar", type="primary", width="stretch", disabled=uploaded_form is None, key=f"btn_parse_form_{case_id}"):
            if uploaded_form is not None:
                with st.spinner("İstem formu okunuyor ve kalemler çıkarılıyor..."):
                    parsed = parse_demand_form(uploaded_form.name, uploaded_form.getvalue(), gemini_key)
                if parsed.get("warning"):
                    st.warning(parsed["warning"])
                new_items = parsed.get("kalemler") or []
                if new_items:
                    update_case_meta(
                        case_id,
                        title=parsed.get("alim_adi") or case_data["title"],
                        procurement_type=parsed.get("alim_turu") or case_data["procurement_type"],
                    )
                    sync_case_items(case_id, new_items)
                    st.success(f"✅ {len(new_items)} kalem çıkarıldı ({parsed.get('source', '')}).")
                    st.rerun()
                else:
                    st.error("Belgeden kalem çıkarılamadı. Lütfen aşağıdaki tabloya elle kalem ekleyin.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_meta:
        st.markdown(
            '<div class="clean-card"><div class="card-title">2. Alım Dosyası Bilgileri</div>'
            '<div class="card-sub">Alım adı ve türünü düzenleyebilir veya doğrudan kalem tablosuna geçebilirsiniz.</div>',
            unsafe_allow_html=True,
        )
        title_val = st.text_input("Alım Adı", value=case_data["title"], key=f"meta_title_{case_id}")
        type_idx = 0 if case_data["procurement_type"] == "Malzeme" else 1
        type_val = st.radio("Alım Türü", ["Malzeme", "Hizmet"], index=type_idx, horizontal=True, key=f"meta_type_{case_id}")
        if st.button("💾 Alım Bilgilerini Güncelle", type="secondary", width="stretch", key=f"btn_save_meta_{case_id}"):
            update_case_meta(case_id, title=title_val.strip() or case_data["title"], procurement_type=type_val)
            st.toast("Dosya bilgileri güncellendi.")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        '<div class="clean-card"><div class="card-title">3. Kalem Listesi (Düzenlenebilir)</div>'
        '<div class="card-sub">Satır ekleyebilir, hücrelere çift tıklayarak düzenleyebilir veya satır seçip Delete tuşuyla silebilirsiniz. İşlem sonrası "Kalemleri Kaydet"e basın.</div>',
        unsafe_allow_html=True,
    )
    if items:
        df_items = pd.DataFrame([
            {
                "id": it["id"],
                "Sıra No": int(it["sira_no"]),
                "Malzeme/Hizmet Adı ve Açıklaması": it["name"],
                "Birim": it["unit"],
                "Miktar": float(it["quantity"]),
            }
            for it in items
        ])
    else:
        df_items = pd.DataFrame([
            {
                "id": None,
                "Sıra No": 1,
                "Malzeme/Hizmet Adı ve Açıklaması": "",
                "Birim": "Adet",
                "Miktar": 1.0,
            }
        ])

    edited_df = st.data_editor(
        df_items,
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "id": None,
            "Sıra No": st.column_config.NumberColumn("Sıra No", min_value=1, step=1, width="small"),
            "Malzeme/Hizmet Adı ve Açıklaması": st.column_config.TextColumn("Malzeme/Hizmet Adı ve Açıklaması", width="large"),
            "Birim": st.column_config.TextColumn("Birim", width="small"),
            "Miktar": st.column_config.NumberColumn("Miktar", min_value=0.01, step=1.0, width="small"),
        },
        key=f"items_editor_{case_id}_{len(items)}",
    )

    c_save1, c_save2 = st.columns([2, 3])
    with c_save1:
        if st.button("💾 Kalemleri Kaydet", type="primary", width="stretch", key=f"btn_save_items_{case_id}"):
            records = edited_df.to_dict(orient="records")
            to_sync: list[dict[str, Any]] = []
            for idx, r in enumerate(records, start=1):
                nm = str(r.get("Malzeme/Hizmet Adı ve Açıklaması") or "").strip()
                if not nm:
                    continue
                sira_raw = r.get("Sıra No")
                sira = int(sira_raw) if (sira_raw is not None and pd.notna(sira_raw)) else idx
                qty_raw = r.get("Miktar")
                qty = float(qty_raw) if (qty_raw is not None and pd.notna(qty_raw) and float(qty_raw) > 0) else 1.0
                row_id = r.get("id")
                to_sync.append({
                    "id": int(row_id) if (row_id is not None and pd.notna(row_id)) else None,
                    "sira_no": sira,
                    "name": nm,
                    "unit": str(r.get("Birim") or "Adet").strip() or "Adet",
                    "quantity": qty,
                })
            sync_case_items(case_id, to_sync)
            st.success(f"✅ {len(to_sync)} kalem kaydedildi.")
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Adım 2: Dosyaya Özel Tedarikçi Arama ve Kayıt
# ---------------------------------------------------------------------------

def render_step2(
    case_data: dict[str, Any],
    items: list[dict[str, Any]],
    suppliers: list[dict[str, Any]],
    gemini_key: str,
) -> None:
    case_id = int(case_data["id"])
    mkys_stats = get_mkys_stats()

    # Genel MKYS Geçmiş Alım Aktarımı (Adım 2 Üst Bölüm)
    with st.expander(
        f"📥 MKYS Geçmiş Alım Excel'i Yükle  ·  (Kayıtlı Geçmiş: {mkys_stats['history_count']} alım · {mkys_stats['unique_suppliers']} firma · Ana Havuz: {mkys_stats['master_count']} firma)",
        expanded=False,
    ):
        st.caption(
            "MKYS İş Zekası'ndan aldığınız genel malzeme/hizmet alım listesini (.xlsx, .xlsm, .csv) yükleyin. "
            "Sistem 'Malzeme/Hizmet Adı' ve 'Tedarikçi Firma Ünvanı' sütunlarını eşleştirip veritabanına kaydeder; "
            "sonraki aramalarda bu firmalar öncelikli getirilerek iletişim bilgileri web'den tamamlanır."
        )
        mc1, mc2 = st.columns([2.0, 1.0])
        with mc1:
            mkys_file = st.file_uploader(
                "MKYS İş Zekası Excel / CSV Dosyası",
                type=["xlsx", "xlsm", "csv"],
                key=f"mkys_uploader_{case_id}",
                label_visibility="collapsed",
            )
        with mc2:
            exclude_pub = st.checkbox("Kamu kurumlarını (DMO vb.) hariç tut", value=True, key=f"mkys_ex_pub_{case_id}")

        if mkys_file is not None:
            raw_mkys = mkys_file.getvalue()
            preview = parse_mkys_excel(mkys_file.name, raw_mkys, exclude_public=exclude_pub)
            if preview.get("error"):
                st.error(preview["error"])
            else:
                hdrs = preview.get("headers") or []
                det_it = preview.get("detected_item_col")
                det_fm = preview.get("detected_firm_col")
                if hdrs:
                    sc1, sc2 = st.columns(2)
                    with sc1:
                        sel_it_col = st.selectbox(
                            "Malzeme / Hizmet Adı Sütunu",
                            options=list(range(len(hdrs))),
                            index=det_it if (det_it is not None and det_it < len(hdrs)) else 0,
                            format_func=lambda i: hdrs[i],
                            key=f"mkys_it_col_{case_id}_{mkys_file.name}",
                        )
                    with sc2:
                        default_fm = det_fm if (det_fm is not None and det_fm < len(hdrs)) else (1 if len(hdrs) > 1 else 0)
                        sel_fm_col = st.selectbox(
                            "Tedarikçi Firma Ünvanı Sütunu",
                            options=list(range(len(hdrs))),
                            index=default_fm,
                            format_func=lambda i: hdrs[i],
                            key=f"mkys_fm_col_{case_id}_{mkys_file.name}",
                        )
                    final_parsed = parse_mkys_excel(
                        mkys_file.name,
                        raw_mkys,
                        exclude_public=exclude_pub,
                        item_col_override=sel_it_col,
                        firm_col_override=sel_fm_col,
                    )
                    recs = final_parsed.get("records") or []
                    st.caption(
                        f"Eşleşen benzersiz alım kaydı: **{len(recs)}** "
                        f"(hariç tutulan kamu kurumu: {final_parsed.get('skipped_public', 0)})"
                    )
                    if st.button("💾 MKYS Geçmiş Alımlarını Ana Havuza Kaydet", type="primary", width="stretch", disabled=not recs, key=f"btn_import_mkys_{case_id}"):
                        for r in recs:
                            if not r.get("sector"):
                                r["sector"] = ", ".join(infer_product_groups(r["item_name"])) or "Genel Tedarik (MKYS)"
                        ins, tot = import_mkys_records(recs, ai_verified=0)
                        if gemini_key:
                            with st.spinner("Yüklenen MKYS kayıtları Gemini AI ile doğrulanıyor ve sektörlendiriliyor..."):
                                unv = get_unverified_pool_records(limit=200)
                                san = sanitize_pool_batch_with_ai(unv["pool"], unv["mkys"], gemini_key, batch_size=20)
                            if san.get("error"):
                                st.error(f"Gemini AI doğrulama hatası (Fail-Closed): {san['error']}")
                            else:
                                res_san = apply_pool_sanitization_results(san["pool_decisions"], san["mkys_decisions"])
                                st.success(
                                    f"✅ {ins} MKYS kaydı işlendi · {res_san['verified']} doğrulanmış firma havuza alındı, "
                                    f"{res_san['deleted']} geçersiz/kamu kaydı elendi."
                                )
                                st.rerun()
                        else:
                            st.error("GEMINI_API_KEY tanımlı değil. Fail-Closed kuralı gereği yüklenen kayıtlar doğrulanamadığı için aramalara dahil edilmedi.")

    if not items:
        st.info("Tedarikçi aramadan önce **1. İstem Formu & Kalemler** adımından en az bir kalem ekleyin.")
        return

    st.markdown(
        '<div class="clean-card"><div class="card-title">Çoklu Motor B2B Tedarikçi Arama (MKYS/Havuz Öncelikli · Yandex · Bing · Yahoo · DDG · İhale)</div>'
        '<div class="card-sub">Önce MKYS geçmiş alımları ve Ana Firma Havuzundaki eşleşmeler getirilir; ardından Gemini sektör tespiti ve B2B sorgularıyla web taranır.</div>',
        unsafe_allow_html=True,
    )

    item_names = [it["name"] for it in items]
    item_by_name = {it["name"]: it for it in items}

    c1, c2, c3 = st.columns([2.2, 1.1, 0.9])
    with c1:
        selected_item_names = st.multiselect(
            "Aranacak Kalemler",
            options=item_names,
            default=item_names,
            key=f"search_items_{case_id}",
        )
    with c2:
        site_limit_opt = st.selectbox(
            "Hedef Site Döngüsü",
            options=[120, 240, 360],
            format_func=lambda n: f"{n} Site ({n // 120} Tur)",
            index=0,
            key=f"site_limit_{case_id}",
        )
    with c3:
        use_gemini = st.checkbox("Gemini B2B Analizi", value=bool(gemini_key), key=f"use_gemini_{case_id}")

    round_state_key = f"case_round_{case_id}"
    visited_state_key = f"case_visited_{case_id}"
    current_round = int(st.session_state.get(round_state_key, 0))
    visited_domains: set[str] = set(st.session_state.get(visited_state_key, set()))

    btn_col1, btn_col2 = st.columns([1.4, 1.0])
    with btn_col1:
        run_search = st.button("🔍 Seçili Kalemler İçin Tedarikçi Ara", type="primary", width="stretch", key=f"btn_search_{case_id}")
    with btn_col2:
        next_round = current_round + 1
        run_next_round = st.button(
            f"🔄 +120 Site Daha Tara ({next_round}. Tur)",
            type="secondary",
            width="stretch",
            disabled=current_round == 0,
            key=f"btn_next_round_{case_id}",
        )

    def _execute_search(start_round: int, end_round: int) -> None:
        if not selected_item_names:
            st.warning("Lütfen en az bir kalem seçin.")
            return
        t0 = time.time()
        cache = load_cache(CACHE_XLSX)
        ai_cache = load_ai_cache()
        progress_bar = st.progress(0)
        status_slot = st.empty()
        total_steps = len(selected_item_names) * (end_round - start_round + 1)
        step_done = 0
        found_total: list[dict[str, Any]] = []
        search_stats: dict[str, Any] = {"inspected_count": 0, "ai_rejected_count": 0, "verified_added_count": 0, "errors": []}

        for idx, it_name in enumerate(selected_item_names, start=1):
            it_obj = item_by_name.get(it_name, {})
            existing_intel = it_obj.get("ai_intel") or {}
            if existing_intel.get("aciklama") and existing_intel.get("b2b_sorgular") and (existing_intel.get("schema_v2") or not use_gemini):
                intel = existing_intel
            else:
                status_slot.info(f"⏳ [{idx}/{len(selected_item_names)}] **{it_name}** analiz ediliyor...")
                intel = analyze_item_with_gemini(it_name, gemini_key if use_gemini else "", ai_cache)
                if it_obj.get("id"):
                    save_item_intel(int(it_obj["id"]), intel)

            for r_num in range(start_round, end_round + 1):
                rows = run_three_ring_search(
                    it_name,
                    intel,
                    cache,
                    progress_cb=lambda m, i=idx: status_slot.info(f"⏳ [{i}/{len(selected_item_names)}] {m}"),
                    max_sites=DEFAULT_MAX_SITES,
                    round_num=r_num,
                    skip_domains=visited_domains,
                    api_key=gemini_key,
                    stats_out=search_stats,
                )
                found_total.extend(rows)
                save_cache(CACHE_XLSX, cache)
                step_done += 1
                progress_bar.progress(int((step_done / max(total_steps, 1)) * 100))

        added_count = add_case_suppliers(case_id, found_total)
        st.session_state[round_state_key] = end_round
        st.session_state[visited_state_key] = visited_domains
        elapsed = round(time.time() - t0, 1)
        st.session_state[f"search_summary_{case_id}"] = {
            "inspected": int(search_stats.get("inspected_count", 0)),
            "rejected": int(search_stats.get("ai_rejected_count", 0)),
            "verified_added": added_count if added_count > 0 else int(search_stats.get("verified_added_count", 0)),
            "elapsed": elapsed,
            "errors": list(search_stats.get("errors") or []),
        }
        st.rerun()

    if run_search:
        target_rounds = max(1, site_limit_opt // DEFAULT_MAX_SITES)
        _execute_search(1, target_rounds)
    elif run_next_round:
        _execute_search(next_round, next_round)

    summary_info = st.session_state.get(f"search_summary_{case_id}")
    if summary_info:
        for err_msg in summary_info.get("errors") or []:
            st.error(f"🚨 {err_msg}")
        st.info(
            f"Tarama tamamlandı: Toplam **{summary_info['inspected']}** site incelendi, "
            f"Gemini AI tarafından **{summary_info['rejected']}** alakasız/spam site elendi, "
            f"**{summary_info['verified_added']}** doğrulanmış tedarikçi eklendi. ({summary_info['elapsed']} sn)"
        )

    st.markdown("</div>", unsafe_allow_html=True)

    # Yapay zeka analiz özeti (varsa)
    analyzed_items = [it for it in items if (it.get("ai_intel") or {}).get("aciklama")]
    if analyzed_items:
        with st.expander("🧠 Yapay Zeka Ürün ve Piyasa Analizleri", expanded=False):
            for it in analyzed_items:
                intel = it["ai_intel"]
                syns = "".join(f'<span class="pill">{html_lib.escape(s)}</span>' for s in intel.get("es_anlamlilar", []))
                brands = "".join(f'<span class="pill">🏷️ {html_lib.escape(b)}</span>' for b in intel.get("markalar", []))
                st.markdown(
                    f'<div class="clean-card" style="padding:12px 16px;"><strong>{html_lib.escape(it["name"])}</strong> '
                    f'<span style="color:#0F766E; font-size:0.8rem;">({html_lib.escape(intel.get("ust_kategori", ""))})</span>'
                    f'<p style="margin:6px 0; color:#334155; font-size:0.88rem;">{html_lib.escape(intel.get("aciklama", ""))}</p>'
                    f'<div>{syns} {brands}</div></div>',
                    unsafe_allow_html=True,
                )

    # Dosyaya Kayıtlı Firmalar Yönetimi
    st.markdown(
        '<div class="clean-card"><div class="card-title">Dosyaya Kayıtlı Tedarikçiler</div>'
        '<div class="card-sub">E-posta listesine dahil edilecek firmaları <b>Seç</b> sütunuyla belirleyin. İstenmeyen firmayı <b>Sil</b> kutucuğuna tıklayarak tek tıkla dosyadan çıkarabilirsiniz.</div>',
        unsafe_allow_html=True,
    )

    all_sectors = sorted({
        sec.strip()
        for s in suppliers
        for sec in str(s.get("sector") or "").split(",")
        if sec.strip()
    })
    tb1, tb1b, tb2, tb3, tb4 = st.columns([1.2, 1.15, 0.85, 0.85, 1.0])
    with tb1:
        filter_q = st.text_input("Firma veya kalem ara...", key=f"sup_filter_{case_id}", label_visibility="collapsed", placeholder="🔍 Firma, e-posta veya kalem...")
    with tb1b:
        sector_filter = st.multiselect(
            "Sektör Filtresi",
            options=all_sectors,
            default=[],
            key=f"sup_sector_filter_{case_id}",
            label_visibility="collapsed",
            placeholder="🏷️ Sektöre göre filtrele...",
        )
    with tb2:
        if st.button("☑️ E-postalıları Seç", width="stretch", key=f"sel_all_mail_{case_id}"):
            edits = [{**s, "selected": bool(s.get("primary_email"))} for s in suppliers]
            update_supplier_selections_and_edits(case_id, edits)
            st.rerun()
    with tb3:
        if st.button("⬜ Seçimi Temizle", width="stretch", key=f"sel_none_mail_{case_id}"):
            edits = [{**s, "selected": False} for s in suppliers]
            update_supplier_selections_and_edits(case_id, edits)
            st.rerun()
    with tb4:
        with st.popover("➕ Manuel Firma Ekle", width="stretch"):
            with st.form(f"manual_sup_form_{case_id}", clear_on_submit=True):
                m_comp = st.text_input("Firma Adı *")
                m_mail = st.text_input("E-posta Adresi")
                m_phone = st.text_input("Telefon")
                m_web = st.text_input("Web Sitesi")
                m_item = st.selectbox("İlgili Kalem", options=item_names)
                m_sec = st.text_input("Sektör", value=(item_by_name.get(m_item, {}).get("ai_intel") or {}).get("ust_kategori", ""))
                if st.form_submit_button("Firmayı Dosyaya Ekle", type="primary", width="stretch"):
                    if m_comp.strip():
                        add_case_suppliers(case_id, [{
                            "item": m_item,
                            "company_name": m_comp.strip(),
                            "primary_email": m_mail.strip().lower(),
                            "all_emails": [m_mail.strip().lower()] if m_mail.strip() else [],
                            "phone": m_phone.strip(),
                            "website": m_web.strip(),
                            "sector": m_sec.strip() or ", ".join(infer_product_groups(m_item)) or "Kurumsal Tedarik",
                            "source": "Manuel Giriş",
                            "match_type": "Manuel",
                            "status": "Aktif",
                        }])
                        st.toast("Firma dosyaya eklendi.")
                        st.rerun()
                    else:
                        st.warning("Firma adı zorunludur.")

    filtered = suppliers
    if filter_q.strip():
        fq = filter_q.strip().lower()
        filtered = [
            s for s in filtered
            if fq in s["company_name"].lower() or fq in s["primary_email"].lower() or fq in s["item_name"].lower() or fq in s["sector"].lower()
        ]
    if sector_filter:
        sel_secs = {x.lower() for x in sector_filter}
        filtered = [
            s for s in filtered
            if any(part.strip().lower() in sel_secs for part in str(s.get("sector") or "").split(",") if part.strip())
        ]

    if not filtered:
        st.info("Bu dosyada henüz kayıtlı tedarikçi bulunmuyor. Yukarıdan arama yapabilir veya manuel firma ekleyebilirsiniz.")
    else:
        df_sup = pd.DataFrame([
            {
                "id": int(s["id"]),
                "Seç": bool(s["selected"]),
                "Sil": False,
                "Firma Adı": s["company_name"],
                "Öncelikli E-posta": s["primary_email"],
                "Telefon": s["phone"],
                "Web Sitesi": s["website"],
                "Kalem": s["item_name"],
                "Sektör": s["sector"],
                "Kaynak": s["source"],
            }
            for s in filtered
        ])

        edited_sup = st.data_editor(
            df_sup,
            width="stretch",
            hide_index=True,
            height=380,
            column_config={
                "id": None,
                "Seç": st.column_config.CheckboxColumn("Seç", width="small"),
                "Sil": st.column_config.CheckboxColumn("Sil", width="small", help="İşaretlenen firma tek tıkla silinir"),
                "Firma Adı": st.column_config.TextColumn("Firma Adı", width="medium"),
                "Öncelikli E-posta": st.column_config.TextColumn("Öncelikli E-posta", width="medium"),
                "Telefon": st.column_config.TextColumn("Telefon", width="small"),
                "Web Sitesi": st.column_config.LinkColumn("Web Sitesi", width="small"),
                "Kalem": st.column_config.TextColumn("Kalem", width="medium"),
                "Sektör": st.column_config.TextColumn("Sektör", width="small"),
                "Kaynak": st.column_config.TextColumn("Kaynak", width="small", disabled=True),
            },
            key=f"sup_editor_{case_id}_{len(suppliers)}_{hash(filter_q)}_{hash(tuple(sector_filter))}",
        )

        # Tek tıkla silme veya düzenleme senkronizasyonu
        edited_records = edited_sup.to_dict(orient="records")
        to_delete_ids = [int(r["id"]) for r in edited_records if r.get("Sil")]
        if to_delete_ids:
            delete_case_suppliers(case_id, to_delete_ids)
            st.toast(f"{len(to_delete_ids)} firma dosyadan silindi.")
            st.rerun()

        # Değişiklik olduysa veritabanını otomatik güncelle
        orig_map = {int(s["id"]): s for s in filtered}
        changed_rows: list[dict[str, Any]] = []
        for r in edited_records:
            sid = int(r["id"])
            orig = orig_map.get(sid)
            if not orig:
                continue
            if (
                bool(r["Seç"]) != bool(orig["selected"])
                or str(r["Firma Adı"] or "") != orig["company_name"]
                or str(r["Öncelikli E-posta"] or "").lower() != orig["primary_email"]
                or str(r["Telefon"] or "") != orig["phone"]
                or str(r["Web Sitesi"] or "") != orig["website"]
                or str(r["Kalem"] or "") != orig["item_name"]
                or str(r["Sektör"] or "") != orig["sector"]
            ):
                changed_rows.append({
                    "id": sid,
                    "selected": bool(r["Seç"]),
                    "company_name": r["Firma Adı"],
                    "primary_email": r["Öncelikli E-posta"],
                    "phone": r["Telefon"],
                    "website": r["Web Sitesi"],
                    "item_name": r["Kalem"],
                    "sector": r["Sektör"],
                })
        if changed_rows:
            update_supplier_selections_and_edits(case_id, changed_rows)
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # Seçili E-postaları Tek Tıkla Kopyalama ve Dışa Aktarma
    selected_emails = collect_bcc_emails(suppliers, only_selected=True)
    st.markdown(
        '<div class="clean-card"><div class="card-title">Seçili Firmaların E-posta Listesi (BCC / Kime)</div>'
        '<div class="card-sub">Yukarıdaki tabloda işaretli firmaların e-posta adresleri aralarına noktalı virgül (;) konarak hazırlanır.</div>',
        unsafe_allow_html=True,
    )
    if selected_emails:
        joined_emails = "; ".join(selected_emails)
        render_copy_button(joined_emails, f"📋 {len(selected_emails)} E-postayı Tek Tıkla Kopyala")
        st.code(joined_emails, language=None)
    else:
        st.info("Seçili firmalara ait e-posta adresi bulunamadı.")

    if suppliers:
        dl1, dl2 = st.columns(2)
        export_df = pd.DataFrame([
            {
                "Kalem": s["item_name"],
                "Firma Adı": s["company_name"],
                "Öncelikli E-posta": s["primary_email"],
                "Tüm E-postalar": "; ".join(s.get("all_emails", [])),
                "Telefon": s["phone"],
                "Web Sitesi": s["website"],
                "Sektör": s["sector"],
                "Kaynak": s["source"],
                "Seçili": "Evet" if s["selected"] else "Hayır",
            }
            for s in suppliers
        ])
        with dl1:
            st.download_button(
                "📥 Firma Listesini İndir (CSV)",
                data=dataframe_to_utf8_bom_csv(export_df),
                file_name=f"tedarikciler_dosya_{case_id}.csv",
                mime="text/csv",
                width="stretch",
            )
        with dl2:
            st.download_button(
                "📊 Firma Listesini İndir (Excel)",
                data=build_excel_bytes(suppliers),
                file_name=f"tedarikciler_dosya_{case_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Adım 3: Resmi Teklife Davet Mektubu (.docx)
# ---------------------------------------------------------------------------

def render_step3(case_data: dict[str, Any], items: list[dict[str, Any]], suppliers: list[dict[str, Any]]) -> None:
    case_id = int(case_data["id"])
    st.markdown(
        '<div class="clean-card"><div class="card-title">Resmi Teklife Davet Mektubu (.docx)</div>'
        '<div class="card-sub">Times New Roman 12 punto, iki yana yaslı resmi yazışma formatında Word belgesi oluşturur.</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1.1, 1.0, 1.0])
    with c1:
        proc_type = st.radio(
            "Alım Türü",
            options=["Malzeme", "Hizmet"],
            index=0 if case_data["procurement_type"] == "Malzeme" else 1,
            horizontal=True,
            key=f"inv_type_{case_id}",
        )
    with c2:
        deadline_date = st.text_input(
            "Son Teklif Tarihi",
            value=case_data.get("deadline_date", ""),
            placeholder="..../..../2026",
            key=f"inv_date_{case_id}",
        )
    with c3:
        deadline_time = st.text_input(
            "Son Teklif Saati",
            value=case_data.get("deadline_time", ""),
            placeholder="...:...",
            key=f"inv_time_{case_id}",
        )

    if (
        proc_type != case_data["procurement_type"]
        or deadline_date != case_data.get("deadline_date", "")
        or deadline_time != case_data.get("deadline_time", "")
    ):
        update_case_meta(
            case_id,
            procurement_type=proc_type,
            deadline_date=deadline_date,
            deadline_time=deadline_time,
        )

    p1_txt, p2_txt = build_invitation_text(deadline_date, deadline_time)
    st.markdown(
        f'<div class="letter-preview"><p style="text-indent:2em; margin-bottom:10px;">{html_lib.escape(p1_txt)}</p>'
        f'<p style="text-indent:2em; margin:0;">{html_lib.escape(p2_txt)}</p></div>',
        unsafe_allow_html=True,
    )

    second_col = "Hizmetin Adı ve Açıklaması" if proc_type == "Hizmet" else "Malzeme Kaleminin Adı"
    if items:
        inv_rows_init = [
            {
                "Sıra No": int(it["sira_no"]),
                second_col: it["name"],
                "Birim": it["unit"],
                "Miktar": format_qty(it["quantity"]),
                "Birim Fiyat": "",
                "Toplam Fiyat": "",
            }
            for it in items
        ]
    else:
        inv_rows_init = [
            {
                "Sıra No": 1,
                second_col: "",
                "Birim": "Hizmet" if proc_type == "Hizmet" else "Adet",
                "Miktar": "",
                "Birim Fiyat": "",
                "Toplam Fiyat": "",
            }
        ]

    st.caption("Word çıktısı öncesi kalem tablosunu buradan da son kez düzenleyebilirsiniz:")
    edited_inv = st.data_editor(
        pd.DataFrame(inv_rows_init),
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        key=f"inv_table_{case_id}_{proc_type}_{len(items)}",
    )

    docx_bytes = build_invitation_docx(
        table_rows=edited_inv.to_dict(orient="records"),
        procurement_type=proc_type,
        deadline_date=deadline_date,
        deadline_time=deadline_time,
        case_title=case_data["title"],
    )
    invited_idents = [
        (s.get("domain") or s.get("company_name") or "").strip()
        for s in suppliers
        if s.get("selected") and (s.get("domain") or s.get("company_name"))
    ]

    def _on_invitation_download() -> None:
        if invited_idents:
            increment_supplier_trust(invited_idents, delta=1)

    st.download_button(
        "📄 Resmi Teklife Davet Mektubunu İndir (.docx)",
        data=docx_bytes,
        file_name=f"teklife_davet_mektubu_{case_id}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        width="stretch",
        on_click=_on_invitation_download,
    )

    with st.expander("✉️ Kısa E-posta Konusu ve Gövde Metni (Kopyala/Yapıştır)", expanded=False):
        subj, body = build_rfq_draft([it["name"] for it in items], deadline_date, deadline_time)
        st.text_input("E-posta Konusu", value=subj, key=f"rfq_subj_{case_id}")
        st.text_area("E-posta Metni", value=body, height=160, key=f"rfq_body_{case_id}")

    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Adım 4: Teklif Değerlendirme, Komisyon İmzaları ve Piyasa Araştırma Tutanağı
# ---------------------------------------------------------------------------

def render_step4(
    case_data: dict[str, Any],
    items: list[dict[str, Any]],
    offers: list[dict[str, Any]],
    gemini_key: str,
) -> None:
    case_id = int(case_data["id"])
    if not items:
        st.info("Teklif değerlendirmesi için önce **1. İstem Formu & Kalemler** adımından kalem ekleyin.")
        return

    # 1. Teklif Girişi: Dijital Yükleme + Manuel Giriş
    col_dig, col_man = st.columns([1.0, 1.25], gap="medium")
    with col_dig:
        st.markdown(
            '<div class="clean-card"><div class="card-title">🤖 Dijital Teklif Yükleme</div>'
            '<div class="card-sub">Maille gelen PDF, Word, Excel veya görsel teklif mektuplarını topluca yükleyin; Gemini firma adını ve birim fiyatları çıkarsın.</div>',
            unsafe_allow_html=True,
        )
        offer_files = st.file_uploader(
            "Teklif Mektupları",
            type=["pdf", "docx", "xlsx", "xlsm", "png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"offer_uploader_{case_id}",
            label_visibility="collapsed",
        )
        if st.button("✨ Teklifleri Yapay Zeka ile Oku ve Ekle", type="primary", width="stretch", disabled=not offer_files, key=f"btn_parse_offers_{case_id}"):
            if offer_files:
                ok_count = 0
                boost_names: list[str] = []
                with st.spinner(f"{len(offer_files)} teklif mektubu okunuyor..."):
                    for uf in offer_files:
                        parsed = parse_offer_document(uf.name, uf.getvalue(), items, gemini_key)
                        f_name = parsed.get("firma_adi") or uf.name
                        upsert_case_offer(
                            case_id=case_id,
                            firm_name=f_name,
                            prices_by_item_id=parsed.get("prices") or {},
                            source_type="Dijital",
                            source_file=uf.name,
                        )
                        boost_names.append(f_name)
                        ok_count += 1
                increment_supplier_trust(boost_names, delta=1)
                st.success(f"✅ {ok_count} dijital teklif tabloya eklendi.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with col_man:
        st.markdown(
            '<div class="clean-card"><div class="card-title">✍️ Manuel Teklif Girişi</div>'
            '<div class="card-sub">Elden veya el yazısıyla gelen teklifler için firma adı ve kalem bazlı birim fiyatları girin.</div>',
            unsafe_allow_html=True,
        )
        man_firm = st.text_input("Teklif Veren Firma Adı", placeholder="Örn: ABC Medikal Ltd. Şti.", key=f"man_firm_{case_id}")
        man_df = pd.DataFrame([
            {
                "item_id": int(it["id"]),
                "Sıra": int(it["sira_no"]),
                "Kalem": it["name"],
                "Miktar": format_qty(it["quantity"]),
                "Birim Fiyat (TL)": 0.0,
            }
            for it in items
        ])
        edited_man = st.data_editor(
            man_df,
            width="stretch",
            hide_index=True,
            height=205,
            column_config={
                "item_id": None,
                "Sıra": st.column_config.NumberColumn("Sıra", disabled=True, width="small"),
                "Kalem": st.column_config.TextColumn("Kalem", disabled=True, width="large"),
                "Miktar": st.column_config.TextColumn("Miktar", disabled=True, width="small"),
                "Birim Fiyat (TL)": st.column_config.NumberColumn("Birim Fiyat (TL)", min_value=0.0, step=0.01, format="%.2f"),
            },
            key=f"man_editor_{case_id}_{len(items)}_{len(offers)}",
        )
        if st.button("💾 Manuel Teklifi Kaydet", type="primary", width="stretch", key=f"btn_save_man_{case_id}"):
            if not man_firm.strip():
                st.warning("Lütfen firma adını girin.")
            else:
                prices: dict[int, float | None] = {}
                for r in edited_man.to_dict(orient="records"):
                    val = parse_tr_number(r.get("Birim Fiyat (TL)"), None)
                    prices[int(r["item_id"])] = val if (val is not None and val > 0) else None
                upsert_case_offer(case_id=case_id, firm_name=man_firm.strip(), prices_by_item_id=prices, source_type="Manuel")
                increment_supplier_trust([man_firm.strip()], delta=1)
                st.success(f"✅ {man_firm.strip()} teklifi kaydedildi.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # 2. Karşılaştırma ve Uhde Hesaplama
    st.markdown(
        '<div class="clean-card"><div class="card-title">Teklif Karşılaştırma ve Uhde Kararı</div>'
        '<div class="card-sub">En uygun teklifler ve uhde kararı otomatik hesaplanarak yeşil renkle vurgulanır.</div>',
        unsafe_allow_html=True,
    )

    if not offers:
        st.info("Karşılaştırma yapabilmek için yukarıdan en az bir dijital veya manuel teklif ekleyin.")
        evaluation = evaluate_case_offers(items, [], case_data.get("evaluation_mode", "Kalem Bazlı En Düşük"))
    else:
        # Kayıtlı teklif yönetimi (düzenleme/silme)
        with st.expander(f"📋 Kayıtlı Teklifleri Yönet ({len(offers)} Firma)", expanded=False):
            for off in offers:
                oc1, oc2, oc3 = st.columns([2.2, 1.2, 0.7])
                quoted = sum(1 for v in off["prices"].values() if v is not None and v > 0)
                oc1.markdown(f"**{off['firm_name']}** · {quoted}/{len(items)} kalem fiyatlı")
                oc2.caption(f"{off['source_type']}" + (f" ({off['source_file']})" if off.get("source_file") else ""))
                if oc3.button("🗑️ Sil", key=f"del_off_{off['id']}", width="stretch"):
                    delete_case_offer(case_id, int(off["id"]))
                    st.toast(f"{off['firm_name']} teklifi silindi.")
                    st.rerun()

        eval_mode = st.radio(
            "Değerlendirme Yöntemi",
            options=["Kalem Bazlı En Düşük", "Toplam Bazlı En Düşük"],
            index=0 if case_data.get("evaluation_mode", "Kalem Bazlı En Düşük") == "Kalem Bazlı En Düşük" else 1,
            horizontal=True,
            key=f"eval_mode_{case_id}",
        )
        if eval_mode != case_data.get("evaluation_mode"):
            update_case_meta(case_id, evaluation_mode=eval_mode)

        evaluation = evaluate_case_offers(items, offers, eval_mode)

        # Görüntülenecek Karşılaştırma DataFrame'i
        comp_rows: list[dict[str, Any]] = []
        win_mask_rows: list[dict[str, bool]] = []
        firm_col_names = [f"{o['firm_name']} (Birim TL)" for o in offers]

        for res in evaluation["item_results"]:
            it = res["item"]
            row_dict: dict[str, Any] = {
                "Sıra": str(it["sira_no"]),
                "Malzeme / Hizmet": it["name"],
                "Birim": it["unit"],
                "Miktar": format_qty(it["quantity"]),
            }
            mask_dict: dict[str, bool] = {}
            winner_oid = res.get("winner_offer_id")
            for o, cname in zip(offers, firm_col_names):
                oid = int(o["id"])
                up = res["firm_prices"].get(oid, {}).get("unit_price")
                row_dict[cname] = f"{format_tr_number(up)} TL" if up else "-"
                mask_dict[cname] = bool(winner_oid is not None and oid == int(winner_oid))

            row_dict["Uhde (Kazanan Firma)"] = res.get("winner_firm") or "-"
            row_dict["En Uygun Toplam"] = f"{format_tr_number(res['winner_total_price'])} TL" if res.get("winner_total_price") else "-"
            mask_dict["Uhde (Kazanan Firma)"] = bool(res.get("winner_firm"))
            mask_dict["En Uygun Toplam"] = bool(res.get("winner_total_price"))
            comp_rows.append(row_dict)
            win_mask_rows.append(mask_dict)

        # Genel Toplam Satırı
        tot_dict: dict[str, Any] = {"Sıra": "", "Malzeme / Hizmet": "GENEL TOPLAM (KDV Hariç)", "Birim": "", "Miktar": ""}
        tot_mask: dict[str, bool] = {}
        tot_winner_id = evaluation.get("total_winner_id")
        for o, cname in zip(offers, firm_col_names):
            oid = int(o["id"])
            ft = evaluation["firm_totals"].get(oid, {})
            tot_dict[cname] = f"{format_tr_number(ft.get('grand_total', 0.0))} TL"
            tot_mask[cname] = bool(tot_winner_id is not None and oid == int(tot_winner_id))
        tot_dict["Uhde (Kazanan Firma)"] = "GENEL UHDE TOPLAMI"
        tot_dict["En Uygun Toplam"] = f"{format_tr_number(evaluation['grand_awarded_total'])} TL"
        tot_mask["Uhde (Kazanan Firma)"] = True
        tot_mask["En Uygun Toplam"] = True
        comp_rows.append(tot_dict)
        win_mask_rows.append(tot_mask)

        comp_df = pd.DataFrame(comp_rows)

        def style_winners(df: pd.DataFrame) -> pd.DataFrame:
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for r_idx, m in enumerate(win_mask_rows):
                is_total_row = r_idx == len(win_mask_rows) - 1
                for col_name in df.columns:
                    if m.get(col_name):
                        styles.loc[r_idx, col_name] = "background-color: #DCFCE7; color: #166534; font-weight: 700;"
                    elif is_total_row:
                        styles.loc[r_idx, col_name] = "background-color: #F1F5F9; font-weight: 700;"
            return styles

        st.dataframe(comp_df.style.apply(style_winners, axis=None), width="stretch", hide_index=True)

        if eval_mode == "Toplam Bazlı En Düşük":
            incomplete = [ft["firm_name"] for ft in evaluation["firm_totals"].values() if not ft["is_complete"]]
            if incomplete:
                st.caption(f"Not: {', '.join(incomplete)} tüm kalemlere fiyat vermediği için tam teklif veren firmalar önceliklendirildi.")

        if evaluation["awarded_summary"]:
            st.markdown("**Uhde Özeti:**")
            u_cols = st.columns(min(len(evaluation["awarded_summary"]) + 1, 4))
            for idx_u, aw in enumerate(evaluation["awarded_summary"][:3]):
                u_cols[idx_u].markdown(
                    f'<div class="kpi-box" style="border-top-color:#16A34A; background:#F0FDF4;">'
                    f'<div class="kpi-lbl">🏆 {html_lib.escape(aw["firm_name"])}</div>'
                    f'<div class="kpi-val" style="font-size:1.15rem; color:#166534;">{format_tr_number(aw["total_amount"])} TL</div>'
                    f'<div style="font-size:0.78rem; color:#15803D;">{aw["item_count"]} kalem uhdesinde</div></div>',
                    unsafe_allow_html=True,
                )
            u_cols[min(len(evaluation["awarded_summary"]), 3)].markdown(
                f'<div class="kpi-box"><div class="kpi-lbl">Genel Uhde Toplamı</div>'
                f'<div class="kpi-val" style="font-size:1.15rem;">{format_tr_number(evaluation["grand_awarded_total"])} TL</div>'
                f'<div style="font-size:0.78rem; color:#64748B;">KDV Hariç · {eval_mode}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)

    # 3. Komisyon ve Harcama Yetkilisi Bloğu
    st.markdown(
        '<div class="clean-card"><div class="card-title">Komisyon Üyeleri ve Hastane Harcama Yetkilisi</div>'
        '<div class="card-sub">En az 1, en fazla 3 komisyon üyesi ile Hastane Harcama Yetkilisini girin. "Varsayılan Olarak Kaydet" ile sonraki dosyalara otomatik aktarılır.</div>',
        unsafe_allow_html=True,
    )
    existing_members = case_data.get("commission") or [{"name": "", "title": "Satın Alma Komisyon Üyesi"}]
    default_cnt = max(1, min(3, len(existing_members)))
    mem_count = st.radio("Komisyon Üye Sayısı", options=[1, 2, 3], index=default_cnt - 1, horizontal=True, key=f"mem_cnt_{case_id}")

    mem_cols = st.columns(mem_count)
    updated_members: list[dict[str, str]] = []
    for i in range(mem_count):
        cur_m = existing_members[i] if i < len(existing_members) else {"name": "", "title": "Satın Alma Komisyon Üyesi"}
        with mem_cols[i]:
            st.markdown(f"**{i + 1}. Komisyon Üyesi**")
            m_name = st.text_input("Adı Soyadı", value=cur_m.get("name", ""), key=f"mem_name_{case_id}_{i}", placeholder="Adı Soyadı")
            m_title = st.text_input("Ünvanı", value=cur_m.get("title", ""), key=f"mem_title_{case_id}_{i}", placeholder="Ünvanı")
            updated_members.append({"name": m_name.strip(), "title": m_title.strip()})

    st.markdown("**Hastane Harcama Yetkilisi (Başhekim)**")
    ac1, ac2 = st.columns(2)
    with ac1:
        auth_name = st.text_input("Başhekim / Harcama Yetkilisi Adı Soyadı", value=case_data.get("authorizer_name", ""), key=f"auth_name_{case_id}")
    with ac2:
        auth_title = st.text_input("Ünvanı", value=case_data.get("authorizer_title", "Başhekim / Hastane Harcama Yetkilisi"), key=f"auth_title_{case_id}")

    cb1, cb2 = st.columns(2)
    with cb1:
        if st.button("💾 Bu Dosyaya Kaydet", type="secondary", width="stretch", key=f"btn_save_comm_{case_id}"):
            update_case_meta(case_id, commission=updated_members, authorizer_name=auth_name, authorizer_title=auth_title)
            st.toast("Komisyon bilgileri bu dosyaya kaydedildi.")
    with cb2:
        if st.button("⭐ Varsayılan Olarak Kaydet", type="primary", width="stretch", key=f"btn_save_def_comm_{case_id}"):
            update_case_meta(case_id, commission=updated_members, authorizer_name=auth_name, authorizer_title=auth_title)
            save_default_commission(updated_members, auth_name, auth_title)
            st.success("✅ Varsayılan komisyon ve harcama yetkilisi kaydedildi. Yeni açılacak tüm dosyalara otomatik gelecektir.")
    st.markdown("</div>", unsafe_allow_html=True)

    # 4. Piyasa Araştırma Tutanağı Word (.docx) Çıktısı
    st.markdown(
        '<div class="clean-card"><div class="card-title">Piyasa Fiyat Araştırma Tutanağı (.docx)</div>'
        '<div class="card-sub">Karşılaştırma tablosu, uhde kararı ve imza bloklarını içeren resmi Word belgesini indirin.</div>',
        unsafe_allow_html=True,
    )
    rc1, rc2 = st.columns([1.0, 2.2])
    with rc1:
        report_date = st.text_input("Tutanak Tarihi", value=dt.date.today().strftime("%d.%m.%Y"), key=f"rep_date_{case_id}")
    with rc2:
        st.markdown("<div style='height: 27px;'></div>", unsafe_allow_html=True)
        tutanak_bytes = build_market_research_docx(
            case_data=case_data,
            offers=offers,
            evaluation=evaluation,
            commission_members=updated_members,
            authorizer_name=auth_name,
            authorizer_title=auth_title,
            report_date=report_date,
        )
        st.download_button(
            "📥 Piyasa Araştırma Tutanağını İndir (.docx)",
            data=tutanak_bytes,
            file_name=f"piyasa_arastirma_tutanagi_{case_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width="stretch",
        )
    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Ana Uygulama Akışı
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def auto_sanitize_pool() -> dict[str, Any]:
    """Uygulama açılışında bir kez arka plan iş parçacığında çalışarak ai_verified = 0 kayıtları sessizce denetler; arayüzü bloklamaz."""
    init_db()
    gemini_key = get_secret("GEMINI_API_KEY", "")

    def _bg_worker() -> None:
        try:
            seed_unverified_from_xlsx(load_cache(CACHE_XLSX))
        except Exception:
            pass
        if not gemini_key:
            return
        for _ in range(15):
            try:
                unv = get_unverified_pool_records(limit=40)
                if not unv["pool"] and not unv["mkys"]:
                    break
                san = sanitize_pool_batch_with_ai(unv["pool"], unv["mkys"], gemini_key, batch_size=20)
                if not san["pool_decisions"] and not san["mkys_decisions"]:
                    break
                apply_pool_sanitization_results(san["pool_decisions"], san["mkys_decisions"])
                if san.get("error"):
                    break
                time.sleep(0.25)
            except Exception:
                break

    t = threading.Thread(target=_bg_worker, name="auto_sanitize_pool", daemon=True)
    t.start()
    return {"started": True}


def main() -> None:
    st.set_page_config(
        page_title="Satın Alma Yönetim Sistemi",
        page_icon="🏥",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_db()
    auto_sanitize_pool()
    inject_minimal_css()
    if not check_optional_login():
        return

    active_case_id = render_sidebar()
    case_data = get_case(active_case_id)
    if case_data is None:
        st.error("Aktif alım dosyası bulunamadı.")
        return

    items = get_case_items(active_case_id)
    suppliers = get_case_suppliers(active_case_id)
    offers = get_case_offers(active_case_id)
    selected_emails = collect_bcc_emails(suppliers, only_selected=True)
    gemini_key = get_secret("GEMINI_API_KEY", "")

    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <h1>🏥 {html_lib.escape(case_data['title'])}</h1>
                <p>Dosya Bazlı Satın Alma Yönetim Sistemi · Çoklu Arama Motoru (120/240/360) &amp; Piyasa Araştırma Süreci</p>
            </div>
            <div class="type-badge">{html_lib.escape(case_data['procurement_type'])} Alımı · Dosya #{active_case_id}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    k1, k2, k3, k4 = st.columns(4)
    for col, lbl, val in [
        (k1, "Kalem Sayısı", len(items)),
        (k2, "Kayıtlı Firma", len(suppliers)),
        (k3, "Seçili E-posta", len(selected_emails)),
        (k4, "Gelen Teklif", len(offers)),
    ]:
        col.markdown(
            f'<div class="kpi-box"><div class="kpi-lbl">{lbl}</div><div class="kpi-val">{val}</div></div>',
            unsafe_allow_html=True,
        )

    tab1, tab2, tab3, tab4 = st.tabs([
        "1️⃣ İstem Formu & Kalemler",
        "2️⃣ Tedarikçi Arama & Kayıt",
        "3️⃣ Teklife Davet Mektubu",
        "4️⃣ Teklif Değerlendirme & Tutanak",
    ])

    with tab1:
        render_step1(case_data, items, gemini_key)
    with tab2:
        render_step2(case_data, items, suppliers, gemini_key)
    with tab3:
        render_step3(case_data, items, suppliers)
    with tab4:
        render_step4(case_data, items, offers, gemini_key)


if __name__ == "__main__":
    main()

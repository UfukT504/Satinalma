"""Kamu Satın Alma | Kurumsal Tedarikçi & Teklif Yönetim Paneli (ERP/SaaS)
Hem yerel bilgisayarda hem de Streamlit Cloud'da çalışır.
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / ".agent" / "skills" / "satinalma-tedarik" / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import parse_form  # noqa: E402
import scrape_suppliers  # noqa: E402

load_dotenv(BASE_DIR / ".env")

CACHE_XLSX = BASE_DIR / "firma_havuzu.xlsx"
OUTPUT_XLSX = BASE_DIR / "tedarikci_listesi.xlsx"
AI_CACHE_JSON = BASE_DIR / "ai_cache.json"

IGNORE_HEADER_WORDS = {
    "SIRA NO",
    "MALZEME ADI",
    "MALZEME ADI VE KISACA AÇIKLAMASI",
    "MALZEME ADI VE KISACA AÇIKMLAMASI",
    "BİRİM",
    "MİKTAR",
    "ADET",
}


def get_secret(key: str, default: str = "") -> str:
    """Bulutta st.secrets üzerinden, yerelde .env / ortam değişkenlerinden güvenli okuma yapar."""
    try:
        if key in st.secrets:
            val = str(st.secrets.get(key, "")).strip()
            if val:
                return val
    except Exception:
        pass
    return (os.getenv(key) or default).strip()


def inject_erp_css() -> None:
    """Kurumsal SaaS/ERP hastane teması (Inter font, #F8FAFC zemin, #0F172A yönetim paneli, #0F766E butonlar)."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        .stApp {
            background-color: #F8FAFC;
            color: #0F172A;
        }
        #MainMenu, footer, header, [data-testid="stToolbar"], .stDeployButton {
            display: none !important;
            visibility: hidden !important;
        }
        .block-container {
            padding-top: 1.2rem !important;
            padding-bottom: 1.5rem !important;
            max-width: 1360px !important;
        }
        .erp-topbar {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            color: #FFFFFF;
            border-radius: 12px;
            padding: 18px 24px;
            margin-bottom: 14px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.14);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .erp-topbar h1 {
            margin: 0;
            font-size: 1.35rem;
            font-weight: 700;
            color: #FFFFFF !important;
            letter-spacing: -0.01em;
        }
        .erp-topbar p {
            margin: 4px 0 0 0;
            font-size: 0.86rem;
            color: #CBD5E1;
        }
        .erp-badge {
            background: rgba(15, 118, 110, 0.28);
            border: 1px solid #14B8A6;
            color: #99F6E4;
            border-radius: 999px;
            padding: 5px 12px;
            font-size: 0.78rem;
            font-weight: 600;
            white-space: nowrap;
        }
        .erp-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 16px 18px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
            margin-bottom: 14px;
        }
        .kpi-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-top: 3px solid #0F766E;
            border-radius: 10px;
            padding: 12px 16px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        }
        .kpi-label {
            font-size: 0.76rem;
            font-weight: 600;
            text-transform: uppercase;
            color: #64748B;
            letter-spacing: 0.03em;
        }
        .kpi-value {
            font-size: 1.45rem;
            font-weight: 700;
            color: #0F172A;
            margin-top: 2px;
        }
        .pill {
            display: inline-block;
            background: #F0FDFA;
            color: #0F766E;
            border: 1px solid #CCFBF1;
            border-radius: 6px;
            padding: 2px 8px;
            font-size: 0.78rem;
            font-weight: 500;
            margin: 2px 4px 2px 0;
        }
        div.stButton > button, div.stDownloadButton > button {
            background-color: #0F766E !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
            padding: 0.5rem 1rem !important;
            transition: background-color 0.15s ease;
        }
        div.stButton > button:hover, div.stDownloadButton > button:hover {
            background-color: #115E59 !important;
            color: #FFFFFF !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_login_screen() -> bool:
    """Açılış şifre ekranı (APP_PASSWORD)."""
    if st.session_state.get("authenticated"):
        return True

    expected_pw = get_secret("APP_PASSWORD", "1234")
    _, mid, _ = st.columns([1.2, 1.1, 1.2])
    with mid:
        st.markdown(
            """
            <div class="erp-topbar" style="display:block; text-align:center;">
                <h1>🏥 Kamu Satın Alma ERP</h1>
                <p>Akıllı Tedarikçi & Teklif Havuzu — Yetkili Erişimi</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            pw = st.text_input("Erişim Şifresi", type="password", placeholder="Şifre (varsayılan: 1234)")
            if st.form_submit_button("Sisteme Giriş Yap", use_container_width=True):
                if pw == expected_pw:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("Hatalı şifre.")
    return False


def load_ai_cache() -> dict[str, dict[str, Any]]:
    """Yerel ai_cache.json önbelleğini yükler."""
    if not AI_CACHE_JSON.exists():
        return {}
    try:
        data = json.loads(AI_CACHE_JSON.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_ai_cache(cache_data: dict[str, dict[str, Any]]) -> None:
    """Gemini ürün analizlerini kota koruması için ai_cache.json dosyasına kaydeder."""
    try:
        AI_CACHE_JSON.write_text(json.dumps(cache_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def analyze_item_with_gemini(
    item_name: str,
    api_key: str,
    ai_cache: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Önce ai_cache.json'a bakar; yoksa Gemini Flash ile 2 cümle tanım, eş anlamlılar, üst kategori ve marka çıkarır."""
    cache_key = re.sub(r"\s+", " ", item_name.strip().lower())
    if cache_key in ai_cache:
        cached = dict(ai_cache[cache_key])
        cached["source"] = f"{cached.get('source', 'ai_cache')} (Önbellek)"
        return cached

    fallback_smart = scrape_suppliers.build_smart_queries(item_name)
    fallback_result = {
        "aciklama": (
            f"{item_name} kalemi hastane birimlerinin doğrudan temin ve piyasa fiyat araştırması kapsamında ihtiyaç duyduğu ürün/hizmettir. "
            "Türkiye'de yetkili medikal, teknik veya endüstriyel tedarikçilerden teklif toplanır."
        ),
        "es_anlamlilar": fallback_smart["queries"][:3],
        "ust_kategori": ", ".join(scrape_suppliers.infer_product_groups(item_name)) or "Kurumsal Tedarik",
        "markalar": [],
        "source": "Yerel Analiz",
    }
    if not api_key:
        return fallback_result

    try:
        from google import genai
        from google.genai import types
    except ImportError:
        return fallback_result

    prompt = (
        f"Kamu hastanesi satın alma birimi için şu kalemi analiz et: '{item_name}'.\n"
        "Yanıtı YALNIZCA aşağıdaki anahtarlara sahip geçerli bir JSON nesnesi olarak ver:\n"
        "- aciklama: Tam olarak 2 cümle; ürünün/hizmetin ne işe yaradığını ve hangi hastane biriminde kullanıldığını açıkla.\n"
        "- es_anlamlilar: Türkiye piyasasında bu kalem için kullanılan 3-5 adet eş anlamlı kelime/arama terimi listesi.\n"
        "- ust_kategori: Bu ürünü satan ana tedarikçi türü (ör. 'tıbbi sarf', 'ecza deposu', 'medikal cihaz distribütörü', 'endüstriyel kimya', 'hastane mobilyası' vb.).\n"
        "- markalar: Türkiye'de bilinen 2-4 üretici veya ana marka adı listesi (yoksa boş liste).\n"
    )
    client = genai.Client(api_key=api_key)
    for model_name in ("gemini-3.5-flash", "gemini-2.5-flash"):
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.2,
                ),
            )
            data = json.loads(resp.text or "{}")
            if data.get("aciklama"):
                result = {
                    "aciklama": str(data.get("aciklama", "")).strip(),
                    "es_anlamlilar": [str(x).strip() for x in (data.get("es_anlamlilar") or []) if str(x).strip()][:5],
                    "ust_kategori": str(data.get("ust_kategori", "")).strip(),
                    "markalar": [str(x).strip() for x in (data.get("markalar") or []) if str(x).strip()][:4],
                    "source": model_name,
                }
                ai_cache[cache_key] = result
                save_ai_cache(ai_cache)
                return result
        except Exception:
            continue
    return fallback_result


def extract_items_from_text(raw_text: str) -> list[str]:
    """Yapıştırılan tablo, virgüllü veya satırlı metinden kalem adlarını çıkarır."""
    items: list[str] = []
    for line in (l.strip() for l in raw_text.splitlines() if l.strip()):
        if "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
            cand = parts[1] if (len(parts) >= 2 and re.fullmatch(r"\d+[.)]?", parts[0])) else parts[0]
            if cand.upper() not in IGNORE_HEADER_WORDS and len(cand) >= 3:
                items.append(cand)
        else:
            for chunk in line.split(","):
                c = re.sub(r"^\d+[.)]\s*", "", chunk.strip())
                if c and c.upper() not in IGNORE_HEADER_WORDS and len(c) >= 3:
                    items.append(c)
    return list(dict.fromkeys(items))


def build_three_ring_task(item_name: str, intel: dict[str, Any]) -> dict[str, Any]:
    """Gemini analizinden 3 halkalı arama görevini (kalem+eş anlamlı + üst kategori + marka) oluşturur."""
    smart = scrape_suppliers.build_smart_queries(item_name)
    synonyms = intel.get("es_anlamlilar") or []
    ust_kat = intel.get("ust_kategori") or ""
    markalar = intel.get("markalar") or []

    ring1 = list(smart["queries"])
    for syn in synonyms[:2]:
        ring1.append(f"{syn} satış iletişim")
    ring1 = list(dict.fromkeys(ring1))[:4]

    ring2 = []
    if ust_kat:
        ring2.append(f"{ust_kat} türkiye tedarikçi iletişim")

    ring3 = [f"{m} türkiye distribütör iletişim" for m in markalar[:2]]

    return {
        "item": item_name,
        "domains": [],
        "queries": ring1,
        "group_queries": ring2,
        "brand_queries": ring3,
        "social_queries": smart["social_queries"],
        "auto_expand": True,
    }


def rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    """Bulunan firma kayıtlarını temiz DataFrame'e dönüştürür."""
    return pd.DataFrame([
        {
            "Kalem": r.get("item", ""),
            "Firma Adı": r.get("company_name", ""),
            "Eşleşme Tipi": r.get("match_type", ""),
            "Öncelikli E-posta": r.get("primary_email", ""),
            "Tüm E-postalar": "; ".join(
                e for e in r.get("all_emails", []) if not str(e).lower().endswith(".kep.tr")
            ),
            "Telefon": r.get("phone", ""),
            "Web Sitesi": r.get("website", ""),
            "Kaynak": r.get("source", ""),
            "Durum": r.get("status", ""),
        }
        for r in rows
    ])


def collect_bcc_emails(rows: list[dict[str, Any]]) -> list[str]:
    """KEP adresleri hariç benzersiz alıcı e-posta listesini döndürür."""
    emails: list[str] = []
    seen: set[str] = set()
    for r in rows:
        candidates = []
        prim = (r.get("primary_email") or "").strip().lower()
        if prim:
            candidates.append(prim)
        for em in r.get("all_emails", []):
            em_l = str(em).strip().lower()
            if em_l and em_l not in candidates:
                candidates.append(em_l)
        for em_l in candidates:
            if not em_l or em_l == "-" or em_l.endswith(".kep.tr") or em_l in seen:
                continue
            seen.add(em_l)
            emails.append(em_l)
    return emails


def dataframe_to_utf8_bom_csv(df: pd.DataFrame) -> bytes:
    """Türkçe Excel uyumlu UTF-8 BOM CSV baytı üretir."""
    buf = io.StringIO()
    df.to_csv(buf, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)
    return buf.getvalue().encode("utf-8-sig")


def build_excel_bytes(rows: list[dict[str, Any]]) -> bytes:
    """İki sekmeli Excel dosyasını bellek içi bayt olarak üretir (Streamlit Cloud tam uyumlu)."""
    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "Tedarikçiler"
    out_rows = [
        [
            r.get("item", ""),
            r.get("company_name", ""),
            r.get("match_type", ""),
            r.get("source", ""),
            r.get("product_groups", ""),
            r.get("website", ""),
            r.get("primary_email", ""),
            "; ".join(e for e in r.get("all_emails", []) if not str(e).lower().endswith(".kep.tr")),
            r.get("phone", ""),
            r.get("status", ""),
            r.get("last_year", "-"),
            r.get("notes", ""),
        ]
        for r in rows
    ]
    scrape_suppliers._write_sheet(ws_main, scrape_suppliers.OUTPUT_COLUMNS, out_rows)
    ws_mails = wb.create_sheet(title="Mail Listesi (Alt Alta)")
    scrape_suppliers._write_sheet(
        ws_mails,
        ["E-posta (Kime/BCC'ye Yapıştır)", "Öncelik", "Kalem", "Firma Adı"],
        scrape_suppliers.collect_unique_email_rows(rows),
    )
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_rfq_draft(items: list[str]) -> tuple[str, str]:
    """4734 sayılı Kanunun 22/d maddesine uygun resmi teklif isteme maili taslağı."""
    items_str = "\n".join(f"- {it}" for it in items)
    subject = f"4734/22-d Doğrudan Temin Piyasa Fiyat Araştırması — {', '.join(items[:2])}"
    body = (
        "Sayın Yetkili,\n\n"
        "Hastanemiz ihtiyacı doğrultusunda 4734 sayılı Kamu İhale Kanunu'nun 22/d maddesi kapsamında "
        "Doğrudan Temin / Piyasa Fiyat Araştırması yapılacaktır.\n\n"
        "Aşağıda belirtilen kalem(ler) için KDV hariç birim fiyat teklifinizi, marka/model bilgisini, "
        "teslim süresini ve varsa ÜTS/teknik doküman bilgilerinizi kaşeli ve imzalı olarak "
        "tarafımıza iletmenizi rica ederiz:\n\n"
        f"{items_str}\n\n"
        "İyi çalışmalar dileriz.\n"
        "Satın Alma Birimi"
    )
    return subject, body


def main() -> None:
    st.set_page_config(
        page_title="Kamu Satın Alma | Akıllı Tedarikçi & Teklif Havuzu",
        page_icon="🏥",
        layout="wide",
    )
    inject_erp_css()

    if not render_login_screen():
        return

    st.markdown(
        """
        <div class="erp-topbar">
            <div>
                <h1>🏥 Kamu Satın Alma | Akıllı Tedarikçi & Teklif Havuzu</h1>
                <p>4734/22-d Piyasa Fiyat Araştırması · Gemini Flash + 3 Halkalı Hızlı Tedarikçi Motoru</p>
            </div>
            <div class="erp-badge">⚡ Hızlı Mod · max 20 site / timeout 3s</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    gemini_key = get_secret("GEMINI_API_KEY", "")

    # Kompakt 2 sütunlu girdi kartı
    st.markdown('<div class="erp-card">', unsafe_allow_html=True)
    col_in, col_file = st.columns([1.45, 1])
    with col_in:
        raw_input = st.text_area(
            "Talep Kalemleri (Her satıra bir kalem yazın veya Excel/Word tablosunu yapıştırın)",
            height=115,
            placeholder="Örn:\nCHİLLER GRUBU İÇİN ANTİFRİZ(ETİLEN GLİKOL)\nDİJİTAL BEBEK BASKÜLÜ",
        )
    with col_file:
        uploaded = st.file_uploader(
            "Talep Formu veya MKYS Dosyası Yükle (.docx, .xlsx, .csv, .pdf)",
            type=["docx", "xlsx", "xlsm", "csv", "pdf", "txt"],
        )
        c_opt1, c_opt2 = st.columns(2)
        with c_opt1:
            use_gemini = st.checkbox("Gemini Flash Analizi", value=bool(gemini_key))
        with c_opt2:
            mode_mkys = st.checkbox("MKYS Firmalarını Tara", value=False)

    btn_col1, btn_col2 = st.columns([4, 1])
    with btn_col1:
        run_clicked = st.button("🔍 Tedarikçileri & Kurumsal Mailleri Bul", use_container_width=True)
    with btn_col2:
        if st.button("🔒 Çıkış", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    if run_clicked:
        t0 = time.time()
        items = extract_items_from_text(raw_input)
        parsed_file: dict[str, Any] = {}
        if uploaded is not None:
            parsed_file = parse_form.extract_items_from_bytes(uploaded.name, uploaded.getvalue())
            if parsed_file.get("error"):
                st.error(parsed_file["error"])
                return
            file_items = [x["item"] for x in parsed_file.get("items", []) if x.get("item")]
            if not items:
                items = file_items or (parsed_file.get("raw_lines") or [])[:15]

        cache = scrape_suppliers.load_cache(CACHE_XLSX)
        ai_cache = load_ai_cache()
        all_rows: list[dict[str, Any]] = []
        intel_map: dict[str, dict[str, Any]] = {}

        progress_bar = st.progress(0)
        status_slot = st.empty()

        if mode_mkys and parsed_file.get("company_details"):
            details = [c for c in parsed_file["company_details"] if not c.get("is_public")]
            item_label = items[0] if items else "MKYS Havuz"
            status_slot.info(f"⏳ MKYS dosyasından {len(details)} firma taranıyor...")
            all_rows = scrape_suppliers.resolve_mkys_companies(
                details,
                item_label,
                cache,
                progress_cb=lambda m: status_slot.info(f"⏳ {m}"),
            )
            progress_bar.progress(100)
        else:
            if not items:
                status_slot.empty()
                st.warning("Lütfen en az bir kalem yazın veya bir dosya yükleyin.")
                return

            total = len(items)
            for idx, item in enumerate(items, start=1):
                status_slot.info(f"⏳ [{idx}/{total}] **{item}** — Yapay zeka analizi yapılıyor...")
                intel = analyze_item_with_gemini(item, gemini_key if use_gemini else "", ai_cache)
                intel_map[item] = intel
                task = build_three_ring_task(item, intel)
                item_rows = scrape_suppliers.process_task(
                    task,
                    cache,
                    progress_cb=lambda m, i=idx: status_slot.info(f"⏳ [{i}/{total}] {m}"),
                    max_sites=20,
                    max_workers=15,
                    page_timeout=3,
                )
                all_rows.extend(item_rows)
                progress_bar.progress(int((idx / total) * 100))

        elapsed = round(time.time() - t0, 1)
        sorted_rows = scrape_suppliers.export_results(OUTPUT_XLSX, CACHE_XLSX, all_rows, cache)
        status_slot.success(f"✅ Tarama {elapsed} saniyede tamamlandı ({len(sorted_rows)} aktif kayıt).")

        st.session_state["last_rows"] = sorted_rows
        st.session_state["last_items"] = items or ["MKYS Havuz"]
        st.session_state["last_intel"] = intel_map
        st.session_state["last_elapsed"] = elapsed

    rows = st.session_state.get("last_rows")
    if rows is None:
        return

    last_items = st.session_state.get("last_items") or []
    intel_map = st.session_state.get("last_intel") or {}
    elapsed_sec = st.session_state.get("last_elapsed", 0.0)

    df = rows_to_dataframe(rows)
    bcc_emails = collect_bcc_emails(rows)

    # 4 adet yatay KPI özet kartı
    k1, k2, k3, k4 = st.columns(4)
    for col, label, val in [
        (k1, "Aranan Kalem", len(last_items)),
        (k2, "Bulunan Firma", len(df)),
        (k3, "Aktif E-Posta", len(bcc_emails)),
        (k4, "İşlem Süresi", f"{elapsed_sec} sn"),
    ]:
        col.markdown(
            f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div></div>',
            unsafe_allow_html=True,
        )

    if intel_map:
        with st.expander("🧠 Yapay Zeka Ürün ve Piyasa Analizleri", expanded=False):
            for item_name, intel in intel_map.items():
                syns = "".join(f'<span class="pill">{s}</span>' for s in intel.get("es_anlamlilar", []))
                brands = "".join(f'<span class="pill">🏷️ {b}</span>' for b in intel.get("markalar", []))
                st.markdown(
                    f"""
                    <div class="erp-card">
                        <strong>{item_name}</strong>
                        <span style="color:#0F766E; font-size:0.82rem;">({intel.get('ust_kategori', '')} · {intel.get('source', '')})</span>
                        <p style="margin:6px 0; color:#334155; font-size:0.9rem;">{intel.get('aciklama', '')}</p>
                        <div>{syns} {brands}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    tab1, tab2, tab3 = st.tabs([
        "🏢 Tedarikçi Tablosu",
        "✉️ Toplu Mail (BCC) & 4734/22-d Taslağı",
        "📥 İndir (CSV & Excel)",
    ])

    with tab1:
        if df.empty:
            st.info("Aktif tedarikçi bulunamadı.")
        else:
            st.dataframe(
                df,
                use_container_width=True,
                height=420,
                hide_index=True,
                column_config={
                    "Web Sitesi": st.column_config.LinkColumn("Web Sitesi"),
                },
            )

    with tab2:
        col_bcc, col_draft = st.columns(2)
        with col_bcc:
            st.markdown("**Tek Tıkla Kopyalanabilir BCC Mail Listesi (KEP Hariç)**")
            if not bcc_emails:
                st.info("Kopyalanacak e-posta bulunamadı.")
            else:
                st.caption("Sağ üstteki kopyala ikonuna basıp Outlook/Zimbra 'Gizli (BCC)' alanına yapıştırın:")
                st.code("; ".join(bcc_emails), language=None)
                st.text_area("Alt Alta Mail Listesi", value="\n".join(bcc_emails), height=150)

        with col_draft:
            st.markdown("**4734 / 22-d Teklif İsteme Maili Taslağı**")
            subj, body = build_rfq_draft(last_items)
            st.text_input("Konu", value=subj)
            st.text_area("Mail Metni", value=body, height=210)

    with tab3:
        if df.empty:
            st.info("İndirilecek veri yok.")
        else:
            c_dl1, c_dl2 = st.columns(2)
            with c_dl1:
                st.download_button(
                    label="📥 UTF-8 BOM CSV İndir (Türkçe Excel Uyumlu)",
                    data=dataframe_to_utf8_bom_csv(df),
                    file_name="tedarikci_listesi.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with c_dl2:
                st.download_button(
                    label="📊 Excel İndir (tedarikci_listesi.xlsx)",
                    data=build_excel_bytes(rows),
                    file_name="tedarikci_listesi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()

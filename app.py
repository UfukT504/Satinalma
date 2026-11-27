"""Kamu Satın Alma | Kurumsal Tedarikçi & Teklif Yönetim Paneli (ERP/SaaS)
Yerel bilgisayarda ve Streamlit Cloud'da ek yerel modüle ihtiyaç duymadan bağımsız çalışır.
"""

from __future__ import annotations

import concurrent.futures
import csv
import datetime as dt
import io
import json
import os
import re
import statistics
import threading
import time
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import openpyxl
import pandas as pd
import requests
import streamlit as st
import urllib3
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openpyxl.styles import Alignment, Font, PatternFill

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

CACHE_XLSX = BASE_DIR / "firma_havuzu.xlsx"
AI_CACHE_JSON = BASE_DIR / "ai_cache.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
}

MARKETPLACE_DOMAINS = {
    "trendyol.com", "hepsiburada.com", "n11.com", "amazon.com.tr", "amazon.com",
    "ciceksepeti.com", "pazarama.com", "pttavm.com", "epttavm.com", "akakce.com",
    "cimri.com", "sahibinden.com", "idefix.com", "alibaba.com", "aliexpress.com",
    "ebay.com", "etsy.com", "wikipedia.org", "youtube.com", "medikalnews.com",
    "medikalfirma.com", "beymen.com", "ebebek.com", "doktortakvimi.com",
    "sikayetvar.com", "yandex.com.tr", "yandex.com", "google.com",
}
SOCIAL_DOMAINS = {"linkedin.com", "instagram.com", "facebook.com", "x.com", "twitter.com"}
TENDER_PORTAL_HINTS = ("ekap.kik.gov.tr", "kik.gov.tr", "ilan.gov.tr", "ihalebul.com", "ihale.com", ".gov.tr", ".bel.tr", ".edu.tr")

MAINTENANCE_KEYWORDS = (
    "bakımda", "bakimda", "yapım aşamasında", "yapim asamasinda",
    "under construction", "coming soon", "alan adı satılık", "alan adı park",
    "park edilmiştir", "park edilmistir", "domain is for sale",
    "this domain may be for sale", "account suspended", "site askıya alındı",
)

EMAIL_RE = re.compile(r"\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,10}\b")
PHONE_RE = re.compile(r"(?:\+90|0)\s*(?:\(?\d{3}\)?[\s.\-]*\d{3}[\s.\-]*\d{2}[\s.\-]*\d{2})")
YEAR_RE = re.compile(r"\b(20[1-3]\d)\b")
COMPANY_NAME_RE = re.compile(
    r"([A-ZÇĞİÖŞÜ][A-Za-zçğıöşüÇĞİÖŞÜ0-9&.\-\s]{2,45}?(?:Medikal|Tıbbi|Sağlık|Ltd\.?\s*Şti\.?|A\.Ş\.?|San\.?\s*Tic\.?))"
)

IGNORED_EMAIL_DOMAINS = {
    "example.com", "sentry.io", "wixpress.com", "wordpress.com",
    "w3.org", "schema.org", "googleapis.com", "googletagmanager.com",
    "yourdomain.com", "domain.com",
}
IGNORED_EMAIL_SUFFIXES = (".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")
VALID_TLDS = {"com", "tr", "net", "org", "info", "biz", "co", "io", "eu", "de", "uk", "gen"}
EMAIL_PRIORITY = ["teklif@", "satis@", "ihale@", "kurumsal@", "info@"]
FREE_EMAIL_PROVIDERS = {"gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "yandex.com"}

CONTACT_PATH_HINTS = (
    "iletisim", "iletişim", "contact", "hakkimizda", "hakkımızda", "kurumsal",
    "bize-ulasin", "bize_ulasin", "ulasin", "ulaşın", "kvkk", "aydinlatma", "aydınlatma", "gizlilik", "kunye", "künye",
)
PRODUCT_PATH_HINTS = ("urun", "ürün", "hizmet", "kategori", "cozum", "çözüm", "marka", "product", "service")

IGNORE_HEADER_WORDS = {
    "SIRA NO", "MALZEME ADI", "MALZEME ADI VE KISACA AÇIKLAMASI",
    "MALZEME ADI VE KISACA AÇIKMLAMASI", "BİRİM", "MİKTAR", "ADET",
}
ITEM_COL_EXCLUDE = ("birim adı", "birim adi", "hastane", "kurum", "bölüm adı", "tarih", "imza", "sıra", "sira")
PUBLIC_ENTITY_RE = re.compile(
    r"\b(DEVLET MALZEME OFİSİ|DMO|MESLEKİ VE TEKNİK|ANADOLU LİSESİ|LİSESİ|SAYMANLIĞI|DÖNER SERMAYE|"
    r"ÜNİVERSİTESİ|HASTANESİ|İL SAĞLIK MÜDÜRLÜĞÜ|BELEDİYESİ|BAKANLIĞI|GENEL MÜDÜRLÜĞÜ)\b",
    re.IGNORECASE,
)
CORP_INDICATOR_RE = re.compile(
    r"\b(LİMİTED|LIMITED|LTD|A\.Ş|ANONİM|ANONIM|SANAYİ|SANAYI|TİCARET|TICARET|ŞTİ|STI|"
    r"MEDİKAL|MEDIKAL|SAĞLIK|SAGLIK|MÜHENDİSLİK|MUHENDISLIK|İNŞAAT|INSAAT|TEKNOLOJİ|TEKNOLOJI|"
    r"YAPI|HIRDAVAT|GRUP|GROUP|TEDARİK|TEDARIK|PAZARLAMA|İTHALAT|ITHALAT|İHRACAT|IHRACAT|"
    r"ELEKTRİK|ELEKTRONİK|REKLAM|TEKSTİL|AMBALAJ|GIDA|OTOMOTİV|MAKİNE|MAKİNA|ENDÜSTRİ|TİC|SAN|ORT)\b",
    re.IGNORECASE,
)
BOILERPLATE_LINE_RE = re.compile(
    r"^(t\.?\s*c\.?|sağlık bakanlığı|devlet hastanesi|başhekim|satın alma|talep formu|"
    r"tarih|sayı|konu|imza|onay|adı soyadı|unvan|birim sorumlusu|not\s*:|açıklama\s*:)",
    re.IGNORECASE,
)

PRODUCT_GROUPS: dict[str, dict[str, Any]] = {
    "Yoğun Bakım & Solunum": {
        "keywords": ("ventilatör", "ventilator", "oksijen", "cpap", "bipap", "spo2", "nabız", "oksimetre", "entübasyon", "aspiratör", "monitör", "defibrilatör", "yoğun bakım", "anestezi"),
        "queries": ["yoğun bakım cihazları medikal distribütör türkiye"],
    },
    "Ameliyathane & Cerrahi": {
        "keywords": ("laparoskop", "trokar", "bistüri", "koter", "cerrahi", "ameliyat", "sütür", "forseps", "endoskop"),
        "queries": ["cerrahi el aletleri ameliyathane medikal tedarikçi türkiye"],
    },
    "Kardiyoloji & EKG": {
        "keywords": ("ekg", "elektrokardiyografi", "holter", "efor", "stent", "anjiyo", "kardiyoloji", "tansiyon"),
        "queries": ["kardiyoloji cihazları ekg medikal distribütör türkiye"],
    },
    "Laboratuvar & Patoloji": {
        "keywords": ("santrifüj", "mikroskop", "otoanalizör", "pcr", "pipet", "reaktif", "laboratuvar", "hematoloji", "biyokimya"),
        "queries": ["tıbbi laboratuvar cihazları distribütör türkiye"],
    },
    "Görüntüleme & Radyoloji": {
        "keywords": ("\\bultrason\\b", "usg", "röntgen", "tomografi", "emg", "radyoloji", "ultrason prob", "mamografi"),
        "queries": ["tıbbi görüntüleme ultrason röntgen distribütör türkiye"],
    },
    "Sterilizasyon & Dezenfeksiyon": {
        "keywords": ("otoklav", "sterilizasyon", "dezenfektan", "plazma", "indikatör", "steril"),
        "queries": ["sterilizasyon otoklav cihazları üretici distribütör türkiye"],
    },
    "Medikal Sarf & Ecza": {
        "keywords": ("enjektör", "eldiven", "kateter", "kanül", "maske", "gazlı bez", "serum", "sonda", "sarf", "ilaç", "ecza", "ampul", "flakon"),
        "queries": ["medikal sarf malzemeleri toptan tedarikçi türkiye", "ecza deposu kurumsal iletişim"],
    },
    "Dental / Diş": {
        "keywords": ("\\bdiş\\b", "dental", "diş üniti", "\\bünit\\b", "türbin", "aerotor", "endodonti", "implant"),
        "queries": ["dental cihaz malzeme deposu türkiye"],
    },
    "Biyomedikal Bakım & Kalibrasyon": {
        "keywords": ("bakım", "onarım", "kalibrasyon", "teknik servis", "periyodik", "tıbbi gaz", "biyomedikal"),
        "queries": ["biyomedikal teknik servis kalibrasyon firmaları türkiye"],
    },
    "Fizik Tedavi & Hastane Demirbaş": {
        "keywords": ("fizik tedavi", "ortopedi", "tekerlekli sandalye", "hasta yatağı", "sedye", "muayene masası", "alt değiştirme", "bakım ünitesi", "hastane mobilyası", "tıbbi mobilya"),
        "queries": ["hastane mobilyaları tıbbi demirbaş üretici"],
    },
    "Pediatri, Bebek & Ölçüm Cihazları": {
        "keywords": ("bebek", "yenidoğan", "pediatri", "baskül", "terazi", "tartı", "boy ve kilo", "boy kilo", "alt değiştirme", "küvöz"),
        "queries": ["medikal bebek baskülü terazi boy ölçer tedarikçi"],
    },
    "Endüstriyel Kimya & Glikol": {
        "keywords": ("antifriz", "glikol", "etilen", "monoetilen", "propilen", "kimyasal", "kazan kimyasal", "su şartlandırma", "klor", "asit"),
        "queries": ["etilen glikol monoetilen glikol tedarikçi satış iletişim"],
    },
    "İklimlendirme, Chiller & HVAC": {
        "keywords": ("chiller", "klima", "soğutma", "fancoil", "kazan", "brülör", "havalandırma", "kompresör"),
        "queries": ["chiller soğutma kimyasalları antifriz tedarikçi iletişim"],
    },
    "Yapı, Hırdavat & İş Güvenliği": {
        "keywords": ("kaydırmaz", "bant", "boya", "kilit", "hırdavat", "el aleti", "iş güvenliği", "kablo", "led"),
        "queries": ["endüstriyel hırdavat iş güvenliği toptan tedarikçi türkiye"],
    },
    "Temizlik, Hijyen & Mutfak": {
        "keywords": ("deterjan", "kağıt havlu", "çöp poşeti", "temizlik", "mutfak", "hijyen"),
        "queries": ["endüstriyel temizlik kimyasalları toptan satış iletişim"],
    },
}
NON_MEDICAL_GROUPS = {"Endüstriyel Kimya & Glikol", "İklimlendirme, Chiller & HVAC", "Yapı, Hırdavat & İş Güvenliği", "Temizlik, Hijyen & Mutfak"}

OUTPUT_COLUMNS = [
    "Kalem", "Firma Adı", "Eşleşme Tipi", "Kaynak", "Ürün Grupları",
    "Web Sitesi", "Öncelikli E-posta", "Tüm E-postalar", "Telefon", "Durum", "Son Güncelleme Yılı", "Notlar",
]
CACHE_COLUMNS = [
    "Firma Adı", "Domain", "Web Sitesi", "Öncelikli E-posta", "Tüm E-postalar",
    "Telefon", "Durum", "Son Güncelleme Yılı", "Son Kontrol", "Notlar", "Ürün Grupları",
]

_THREAD_LOCAL = threading.local()


# ---------------------------------------------------------------------------
# Yapılandırma ve Arayüz Stilleri
# ---------------------------------------------------------------------------

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
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }
        .stApp { background-color: #F8FAFC; color: #0F172A; }
        #MainMenu, footer, header, [data-testid="stToolbar"], .stDeployButton { display: none !important; visibility: hidden !important; }
        .block-container { padding-top: 1.2rem !important; padding-bottom: 1.5rem !important; max-width: 1360px !important; }
        .erp-topbar {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            color: #FFFFFF; border-radius: 12px; padding: 18px 24px; margin-bottom: 14px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.14); display: flex; justify-content: space-between; align-items: center;
        }
        .erp-topbar h1 { margin: 0; font-size: 1.35rem; font-weight: 700; color: #FFFFFF !important; }
        .erp-topbar p { margin: 4px 0 0 0; font-size: 0.86rem; color: #CBD5E1; }
        .erp-badge {
            background: rgba(15, 118, 110, 0.28); border: 1px solid #14B8A6; color: #99F6E4;
            border-radius: 999px; padding: 5px 12px; font-size: 0.78rem; font-weight: 600; white-space: nowrap;
        }
        .erp-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
            padding: 16px 18px; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04); margin-bottom: 14px;
        }
        .kpi-card {
            background: #FFFFFF; border: 1px solid #E2E8F0; border-top: 3px solid #0F766E;
            border-radius: 10px; padding: 12px 16px; box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        }
        .kpi-label { font-size: 0.76rem; font-weight: 600; text-transform: uppercase; color: #64748B; letter-spacing: 0.03em; }
        .kpi-value { font-size: 1.45rem; font-weight: 700; color: #0F172A; margin-top: 2px; }
        .pill {
            display: inline-block; background: #F0FDFA; color: #0F766E; border: 1px solid #CCFBF1;
            border-radius: 6px; padding: 2px 8px; font-size: 0.78rem; font-weight: 500; margin: 2px 4px 2px 0;
        }
        div.stButton > button, div.stDownloadButton > button {
            background-color: #0F766E !important; color: #FFFFFF !important; border: none !important;
            border-radius: 8px !important; font-weight: 600 !important; font-size: 0.9rem !important; padding: 0.5rem 1rem !important;
        }
        div.stButton > button:hover, div.stDownloadButton > button:hover { background-color: #115E59 !important; color: #FFFFFF !important; }
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
            '<div class="erp-topbar" style="display:block; text-align:center;">'
            '<h1>🏥 Kamu Satın Alma ERP</h1><p>Akıllı Tedarikçi & Teklif Havuzu — Yetkili Erişimi</p></div>',
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


# ---------------------------------------------------------------------------
# Dosya Ayrıştırma (Word, Excel, CSV, PDF, Metin — Bağımsız)
# ---------------------------------------------------------------------------

def _tr_lower(text: str) -> str:
    return text.replace("İ", "i").replace("I", "ı").lower()


def _tr_upper(text: str) -> str:
    return text.replace("i", "İ").replace("ı", "I").upper()


def _clean_item_text(text: str) -> str:
    cleaned = re.sub(r"^\s*\d+[.)\-]\s*", "", text).strip()
    cleaned = re.sub(r"^[•\-–*]\s*", "", cleaned).strip()
    cleaned = re.sub(r"\s*[-–(]?\s*\b\d+\s*(?:adet|kutu|paket|set|takım|kg|lt|metre|m)\b\)?\s*$", "", cleaned, flags=re.IGNORECASE).strip()
    return re.sub(r"\s+", " ", cleaned)


def _detect_columns(headers: list[str]) -> dict[str, int]:
    header_keywords = {
        "item": ["malzeme tanımı", "malzeme tanimi", "malzeme adı", "cihaz adı", "ürün adı", "hizmet adı", "kalem adı", "malzeme", "cihaz", "kalem", "ürün", "hizmet", "cinsi", "açıklama"],
        "company": ["tedarikçi firma", "tedarikci firma", "firma", "tedarikçi", "yüklenici", "satıcı", "unvan", "ünvan"],
    }
    mapping: dict[str, int] = {}
    lowered = [str(h or "").strip().lower() for h in headers]
    for role, keys in header_keywords.items():
        for k in keys:
            for idx, col in enumerate(lowered):
                if not col:
                    continue
                if role == "item" and any(ex in col for ex in ITEM_COL_EXCLUDE):
                    continue
                if k in col:
                    mapping[role] = idx
                    break
            if role in mapping:
                break
    return mapping


def _classify_supplier(name: str) -> dict[str, Any]:
    raw = re.sub(r"\s+", " ", str(name)).strip()
    is_pub = bool(PUBLIC_ENTITY_RE.search(_tr_upper(raw)))
    is_corp = bool(CORP_INDICATOR_RE.search(raw))
    words = [w for w in re.split(r"[\s\-]+", raw) if w]
    is_person = (not is_corp) and (not is_pub) and (2 <= len(words) <= 4)
    stop_tokens = {"SANAYİ", "SANAYI", "TİCARET", "TICARET", "LİMİTED", "LIMITED", "ŞİRKETİ", "LTD", "ŞTİ", "A.Ş", "VE", "İTHALAT", "İHRACAT", "PAZARLAMA", "SAN", "TİC", "ORT"}
    core_words: list[str] = []
    for w in words:
        w_clean = re.sub(r"[^\wçğıöşüÇĞİÖŞÜ]", "", _tr_upper(w))
        if not w_clean:
            continue
        if is_corp and w_clean in stop_tokens and len(core_words) >= 2:
            break
        core_words.append(w)
        if is_corp and len(core_words) >= 3:
            break
    return {
        "name": raw,
        "type": "Kamu / DMO" if is_pub else ("Şahıs Firması" if is_person else "Kurumsal"),
        "is_public": is_pub,
        "is_person": is_person,
        "search_core": " ".join(core_words) or raw,
    }


def _analyze_table(rows: list[list[Any]]) -> dict[str, Any]:
    non_empty = [r for r in rows if any(str(c or "").strip() for c in r)]
    if not non_empty:
        return {"items": [], "companies": [], "company_details": [], "raw_lines": []}
    best_idx, best_score, best_map = 0, -1, {}
    for i, r in enumerate(non_empty[:100]):
        if sum(1 for c in r if str(c or "").strip()) < 2:
            continue
        cand_map = _detect_columns([str(c or "").strip() for c in r])
        score = len(cand_map) + (3 if "company" in cand_map else 0) + (2 if "item" in cand_map else 0)
        if score > best_score:
            best_idx, best_score, best_map = i, score, cand_map

    item_col = best_map.get("item")
    comp_col = best_map.get("company")
    items_seen: dict[str, dict[str, Any]] = {}
    companies_seen: dict[str, str] = {}
    for r in non_empty[best_idx + 1:]:
        it_val = str(r[item_col]).strip() if (item_col is not None and item_col < len(r) and r[item_col] is not None) else ""
        co_val = str(r[comp_col]).strip() if (comp_col is not None and comp_col < len(r) and r[comp_col] is not None) else ""
        if co_val and len(co_val) > 2:
            companies_seen.setdefault(_tr_upper(co_val), co_val)
        if it_val and len(it_val) > 1:
            items_seen.setdefault(it_val, {"item": it_val})

    raw_lines = [" | ".join(str(c).strip() for c in r if str(c or "").strip()) for r in non_empty[:100]] if not items_seen else []
    comps = sorted(companies_seen.values())
    return {
        "detected_columns": best_map,
        "items": list(items_seen.values()),
        "companies": comps,
        "company_details": [_classify_supplier(c) for c in comps],
        "raw_lines": raw_lines,
    }


def parse_uploaded_bytes(filename: str, raw_bytes: bytes) -> dict[str, Any]:
    """Yüklenen .docx, .xlsx, .csv, .pdf, .txt dosyasından kalemleri ve MKYS firmalarını çıkarır."""
    ext = Path(filename).suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
            rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
            wb.close()
            return _analyze_table(rows)
        if ext == ".csv":
            text = raw_bytes.decode("utf-8-sig", errors="replace")
            lines = text.splitlines()
            sample = "\n".join(lines[:30])
            delim = ";" if sample.count(";") > sample.count(",") else ","
            return _analyze_table(list(csv.reader(lines, delimiter=delim)))
        if ext == ".docx":
            import docx

            doc = docx.Document(io.BytesIO(raw_bytes))
            merged_items: list[dict[str, Any]] = []
            seen_keys: set[str] = set()
            companies: set[str] = set()
            for tbl in doc.tables:
                tbl_rows = [[cell.text.strip().replace("\n", " ") for cell in row.cells] for row in tbl.rows]
                res = _analyze_table(tbl_rows)
                if "item" not in res.get("detected_columns", {}):
                    continue
                for it in res["items"]:
                    name = _clean_item_text(it["item"])
                    if name and len(name) >= 3 and not BOILERPLATE_LINE_RE.match(name) and name.lower() not in seen_keys:
                        seen_keys.add(name.lower())
                        merged_items.append({"item": name})
                companies.update(res.get("companies", []))
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            if not merged_items:
                for p in paragraphs:
                    if not BOILERPLATE_LINE_RE.match(p) and re.match(r"^\s*(?:\d+[.)\-]|[•\-–*])\s+", p):
                        cand = _clean_item_text(p)
                        if cand and len(cand) >= 3 and cand.lower() not in seen_keys:
                            seen_keys.add(cand.lower())
                            merged_items.append({"item": cand})
            comps = sorted(companies)
            return {
                "items": merged_items,
                "companies": comps,
                "company_details": [_classify_supplier(c) for c in comps],
                "raw_lines": [p for p in paragraphs if not BOILERPLATE_LINE_RE.match(p)][:150],
            }
        if ext == ".pdf":
            from pypdf import PdfReader

            reader = PdfReader(io.BytesIO(raw_bytes))
            lines = [l.strip() for p in reader.pages for l in (p.extract_text() or "").splitlines() if l.strip()]
            return {"items": [], "companies": [], "company_details": [], "raw_lines": lines[:200]}
        if ext == ".txt":
            lines = [l.strip() for l in raw_bytes.decode("utf-8-sig", errors="replace").splitlines() if l.strip()]
            return {"items": [], "companies": [], "company_details": [], "raw_lines": lines[:200]}
    except Exception as exc:
        return {"error": f"Dosya okuma hatası: {exc}", "items": [], "companies": [], "company_details": []}
    return {"error": f"Desteklenmeyen dosya türü: {ext}", "items": [], "companies": [], "company_details": []}


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


# ---------------------------------------------------------------------------
# Gemini Flash Analizi + ai_cache.json Kota Koruması
# ---------------------------------------------------------------------------

def infer_product_groups(text: str) -> list[str]:
    low = text.lower()
    matched: list[str] = []
    for group_name, cfg in PRODUCT_GROUPS.items():
        for kw in cfg["keywords"]:
            if (kw.startswith("\\b") and re.search(kw, low)) or (not kw.startswith("\\b") and kw in low):
                matched.append(group_name)
                break
    return matched


def build_smart_queries(item_name: str) -> dict[str, Any]:
    raw = str(item_name).strip()
    paren_parts = [p.strip() for p in re.findall(r"\(([^)]+)\)", raw) if len(p.strip()) >= 3]
    outside = re.sub(r"\([^)]*\)", " ", raw)
    outside_clean = re.sub(r"\s+", " ", re.sub(r"\b(için|icin|grubu|tipi|marka|model|adet)\b", " ", outside, flags=re.IGNORECASE)).strip()
    terms = [t for t in ([outside_clean] + paren_parts) if t]
    if any("etilen" in _tr_lower(p) and "glikol" in _tr_lower(p) for p in paren_parts):
        terms.append("monoetilen glikol")
    inferred = infer_product_groups(" ".join([raw] + terms))
    is_non_med = any(g in NON_MEDICAL_GROUPS for g in inferred) and not any(g not in NON_MEDICAL_GROUPS for g in inferred)
    queries = []
    for t in list(dict.fromkeys(terms))[:3]:
        queries.append(f"{t} {'satış iletişim' if is_non_med else 'medikal iletişim'}")
        queries.append(f"{t} {'toptan tedarikçi' if is_non_med else 'tedarikçi satış'}")
    return {"queries": list(dict.fromkeys(queries))[:4], "social_queries": [terms[0] if terms else raw]}


def load_ai_cache() -> dict[str, dict[str, Any]]:
    if not AI_CACHE_JSON.exists():
        return {}
    try:
        data = json.loads(AI_CACHE_JSON.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_ai_cache(cache_data: dict[str, dict[str, Any]]) -> None:
    try:
        AI_CACHE_JSON.write_text(json.dumps(cache_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def analyze_item_with_gemini(item_name: str, api_key: str, ai_cache: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Önce ai_cache.json'a bakar; yoksa Gemini Flash ile tanım, eş anlamlılar, üst kategori ve marka çıkarır."""
    cache_key = re.sub(r"\s+", " ", item_name.strip().lower())
    if cache_key in ai_cache:
        cached = dict(ai_cache[cache_key])
        cached["source"] = f"{cached.get('source', 'ai_cache')} (Önbellek)"
        return cached

    fallback_smart = build_smart_queries(item_name)
    fallback_result = {
        "aciklama": (
            f"{item_name} kalemi hastane birimlerinin doğrudan temin ve piyasa fiyat araştırması kapsamında ihtiyaç duyduğu ürün/hizmettir. "
            "Türkiye'de yetkili medikal, teknik veya endüstriyel tedarikçilerden teklif toplanır."
        ),
        "es_anlamlilar": fallback_smart["queries"][:3],
        "ust_kategori": ", ".join(infer_product_groups(item_name)) or "Kurumsal Tedarik",
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
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2),
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


# ---------------------------------------------------------------------------
# Tarayıcısız 3 Halkalı Arama & İletişim Kazıma Motoru (timeout=3, verify=False)
# ---------------------------------------------------------------------------

def _root_domain(url_or_domain: str) -> str:
    raw = url_or_domain.strip().lower()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "https://" + raw
    host = urlparse(raw).netloc.split(":")[0]
    return host[4:] if host.startswith("www.") else host


def _is_excluded_domain(domain: str) -> bool:
    return any(domain == m or domain.endswith("." + m) for m in MARKETPLACE_DOMAINS)


def _is_social(domain: str) -> bool:
    return any(domain == s or domain.endswith("." + s) for s in SOCIAL_DOMAINS)


def _is_tender_domain(domain: str) -> bool:
    return any(domain == h or domain.endswith(h) or h in domain for h in TENDER_PORTAL_HINTS)


def rank_emails(emails: set[str], site_domain: str) -> tuple[str, list[str]]:
    """KEP adreslerini eler ve teklif@/satis@/ihale@/kurumsal@/info@ sırasına göre sıralar."""
    cleaned: list[str] = []
    for e in emails:
        el = e.lower().strip(".")
        if el.endswith(IGNORED_EMAIL_SUFFIXES) or el.startswith("www."):
            continue
        dom = el.split("@")[-1]
        tld = dom.rsplit(".", 1)[-1] if "." in dom else ""
        if dom in IGNORED_EMAIL_DOMAINS or dom == "kep.tr" or dom.endswith(".kep.tr") or tld not in VALID_TLDS:
            continue
        cleaned.append(el)
    unique = sorted(set(cleaned))
    if not unique:
        return "", []

    def score(email: str) -> tuple[int, int, str]:
        prefix_rank = next((i for i, p in enumerate(EMAIL_PRIORITY) if email.startswith(p)), len(EMAIL_PRIORITY))
        dom = email.split("@")[-1]
        is_corp = 0 if (site_domain and site_domain in dom) else (2 if dom in FREE_EMAIL_PROVIDERS else 1)
        return (prefix_rank, is_corp, email)

    ordered = sorted(unique, key=score)
    return ordered[0], ordered


def _get_session() -> requests.Session:
    sess = getattr(_THREAD_LOCAL, "session", None)
    if sess is None:
        sess = requests.Session()
        sess.headers.update(HEADERS)
        _THREAD_LOCAL.session = sess
    return sess


def fetch_page(url: str, timeout: int = 3) -> tuple[int | None, str, str]:
    """Tarayıcı açmadan timeout=3 ve verify=False ile sayfa getirir."""
    try:
        resp = _get_session().get(url, timeout=timeout, verify=False, allow_redirects=True)
        resp.encoding = resp.apparent_encoding or "utf-8"
        return resp.status_code, resp.url, resp.text
    except requests.RequestException:
        return None, url, ""


def _extract_contacts(html: str, base_url: str) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    emails: set[str] = set()
    contact_links: list[str] = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        hl = href.lower()
        at = a.get_text(" ", strip=True).lower()
        if hl.startswith("mailto:"):
            addr = href[7:].split("?")[0].strip()
            if EMAIL_RE.match(addr):
                emails.add(addr.lower())
        elif any(h in hl or h in at for h in CONTACT_PATH_HINTS):
            full = urljoin(base_url, href)
            if urlparse(full).scheme in ("http", "https"):
                contact_links.append(full)
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    emails.update(m.lower() for m in EMAIL_RE.findall(text))
    phones = list(dict.fromkeys(re.sub(r"\s+", " ", re.sub(r"[()]", "", p)).strip() for p in PHONE_RE.findall(text)))
    current_year = dt.date.today().year
    years = [int(y) for y in YEAR_RE.findall(text) if 2010 <= int(y) <= current_year + 1]
    is_maint = any(k in (title + " " + text[:4000]).lower() for k in MAINTENANCE_KEYWORDS)
    return {
        "title": title, "emails": emails, "phones": phones, "max_year": max(years) if years else None,
        "is_maintenance": is_maint, "contact_links": list(dict.fromkeys(contact_links))[:4],
        "groups": infer_product_groups(title + " " + text[:10000]),
    }


def inspect_company_website(dom: str, company_hint: str = "", ignore_year_filter: bool = False, page_timeout: int = 3) -> dict[str, Any]:
    """Siteyi ve /iletisim, /contact, /kurumsal sayfalarını tarayarak aktif mail/telefon çıkarır."""
    status, final_url, html = fetch_page(f"https://{dom}", timeout=page_timeout)
    if status is None:
        status, final_url, html = fetch_page(f"http://{dom}", timeout=page_timeout)
    if status is None or status >= 400 or not html.strip():
        return {"status": "Pasif"}

    home = _extract_contacts(html, final_url)
    final_dom = _root_domain(final_url)
    if not final_dom.endswith(".tr"):
        tr_hints = ("türkiye", "turkiye", "istanbul", "ankara", "izmir", "bursa", "iletisim", "iletişim", "ticaret", "sanayi", "+90")
        if not any(h in html[:25000].lower() for h in tr_hints):
            return {"status": "Elendi (Yurt Dışı)"}
    if home["is_maintenance"]:
        return {"status": "Pasif (Bakımda)"}

    all_emails = set(home["emails"])
    phones = list(home["phones"])
    years = [home["max_year"]] if home["max_year"] else []
    groups = set(home["groups"])

    cand_links = list(home["contact_links"])
    for fb in ("/iletisim", "/contact", "/kurumsal"):
        u = urljoin(final_url, fb)
        if u not in cand_links:
            cand_links.append(u)

    home_prim, _ = rank_emails(all_emails, final_dom)
    max_sub = 1 if (home_prim and phones) else 2
    for link in cand_links[:max_sub]:
        if _root_domain(link) != final_dom:
            continue
        cs, cu, ch = fetch_page(link, timeout=page_timeout)
        if cs and cs < 400 and ch:
            cd = _extract_contacts(ch, cu)
            all_emails.update(cd["emails"])
            for p in cd["phones"]:
                if p not in phones:
                    phones.append(p)
            if cd["max_year"]:
                years.append(cd["max_year"])
            groups.update(cd["groups"])

    max_year = max(years) if years else None
    if max_year is not None and max_year < 2023 and not ignore_year_filter:
        return {"status": f"Pasif ({max_year})"}

    prim, ordered = rank_emails(all_emails, final_dom)
    return {
        "company_name": (company_hint or home["title"].split("|")[0].split("-")[0].strip() or final_dom)[:80],
        "domain": final_dom,
        "website": final_url,
        "primary_email": prim,
        "all_emails": ordered,
        "phone": " / ".join(phones[:2]),
        "status": "Aktif",
        "last_year": max_year or "-",
        "notes": "" if prim else "Web form üzerinden iletişim",
        "product_groups": ", ".join(sorted(groups)),
    }


def _search_backend(query: str, backend: str | None, max_results: int) -> list[dict[str, str]]:
    ddgs_cls = None
    is_legacy = False
    try:
        from ddgs import DDGS
        ddgs_cls = DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
            ddgs_cls = DDGS
            is_legacy = True
        except ImportError:
            return []
    try:
        client = ddgs_cls(timeout=3)
        kwargs: dict[str, Any] = {"region": "tr-tr", "max_results": max_results}
        if not is_legacy and backend:
            kwargs["backend"] = backend
        return [{"title": r.get("title", ""), "href": r.get("href", ""), "body": r.get("body", "")} for r in client.text(query, **kwargs)]
    except Exception:
        return []


def search_multi_engine(query: str, max_results: int = 12) -> list[dict[str, str]]:
    merged: list[dict[str, str]] = []
    seen: set[str] = set()
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
        futs = [pool.submit(_search_backend, query, b, max_results) for b in ("google", "bing", "yahoo")]
        for fut in concurrent.futures.as_completed(futs):
            for r in fut.result():
                href = (r.get("href") or "").strip()
                if href and href not in seen:
                    seen.add(href)
                    merged.append(r)
    return merged


def load_cache(cache_path: Path) -> dict[str, dict[str, Any]]:
    cache: dict[str, dict[str, Any]] = {}
    if not cache_path.exists():
        return cache
    try:
        wb = openpyxl.load_workbook(cache_path, data_only=True)
        rows = list(wb.active.iter_rows(values_only=True))
        wb.close()
        for r in rows[1:]:
            if not r or not r[1]:
                continue
            dom = str(r[1]).strip().lower()
            prim = str(r[3] or "").strip().lower()
            if prim.endswith(".kep.tr"):
                prim = ""
            all_em = [x.strip().lower() for x in re.split(r"[,;]", str(r[4] or "")) if x.strip() and not x.strip().lower().endswith(".kep.tr")]
            cache[dom] = {
                "company_name": str(r[0] or dom), "domain": dom, "website": str(r[2] or ""),
                "primary_email": prim, "all_emails": all_em, "phone": str(r[5] or ""),
                "status": str(r[6] or ""), "last_year": r[7] or "-", "last_checked": str(r[8] or ""),
                "notes": str(r[9] or ""), "product_groups": str(r[10] or "") if len(r) > 10 else "",
            }
    except Exception:
        pass
    return cache


def run_three_ring_search(
    item_name: str,
    intel: dict[str, Any],
    cache: dict[str, dict[str, Any]],
    progress_cb: Any = None,
    max_sites: int = 20,
) -> list[dict[str, Any]]:
    """3 halkalı aramayı (kalem+eş anlamlı + üst kategori + marka/EKAP) 15 iş parçacığı ve timeout=3 ile çalıştırır."""
    smart = build_smart_queries(item_name)
    ring1 = list(dict.fromkeys(smart["queries"] + [f"{s} satış iletişim" for s in (intel.get("es_anlamlilar") or [])[:2]]))[:4]
    ust_kat = intel.get("ust_kategori") or ""
    inferred = infer_product_groups(f"{item_name} {ust_kat}")
    ring2 = [f"{ust_kat} türkiye tedarikçi iletişim"] if ust_kat else []
    for g in inferred[:2]:
        ring2.extend(PRODUCT_GROUPS[g]["queries"][:1])
    ring3 = [f"{m} türkiye distribütör iletişim" for m in (intel.get("markalar") or [])[:2]]
    ring3.append(f"{item_name} doğrudan temin")

    query_specs: list[tuple[str, str, str, int]] = (
        [(q, "Web Arama", "1. Halka (Doğrudan)", 12) for q in ring1]
        + [(q, "Web Arama (Kategori)", "2. Halka (Üst Kategori)", 10) for q in list(dict.fromkeys(ring2))[:3]]
        + [(q, "Marka / İhale İzi", "3. Halka (Marka / EKAP)", 8) for q in ring3]
    )

    if progress_cb:
        progress_cb(f"{item_name}: {len(query_specs)} sorgu taranıyor...")

    candidate_domains: dict[str, tuple[str, str]] = {}
    snippet_fallbacks: dict[str, dict[str, Any]] = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as qpool:
        q_futs = {qpool.submit(search_multi_engine, q, mr): (sl, mt) for q, sl, mt, mr in query_specs}
        for qfut in concurrent.futures.as_completed(q_futs):
            sl, mt = q_futs[qfut]
            for res in qfut.result():
                href = res.get("href", "")
                dom = _root_domain(href)
                if not dom or _is_excluded_domain(dom) or _is_social(dom) or _is_tender_domain(dom):
                    continue
                candidate_domains.setdefault(dom, (sl, mt))
                txt = f"{res.get('title', '')} {res.get('body', '')}"
                s_em = set(EMAIL_RE.findall(txt))
                s_ph = [re.sub(r"\s+", " ", re.sub(r"[()]", "", p)).strip() for p in PHONE_RE.findall(txt)]
                if s_em or s_ph:
                    fb = snippet_fallbacks.setdefault(dom, {"emails": set(), "phones": [], "title": res.get("title", ""), "url": href})
                    fb["emails"].update(s_em)
                    fb["phones"].extend(s_ph)

    results: list[dict[str, Any]] = []
    to_fetch: list[tuple[str, str, str]] = []
    for dom, (sl, mt) in candidate_domains.items():
        if dom in cache and cache[dom].get("status") == "Aktif" and cache[dom].get("primary_email"):
            hit = dict(cache[dom])
            hit.update({"item": item_name, "source": f"{sl} (Önbellek)", "match_type": mt})
            results.append(hit)
        else:
            to_fetch.append((dom, sl, mt))

    to_fetch = to_fetch[:max_sites]
    if to_fetch:
        total = len(to_fetch)
        with concurrent.futures.ThreadPoolExecutor(max_workers=15) as pool:
            futs = {pool.submit(inspect_company_website, dom, "", False, 3): (dom, sl, mt) for dom, sl, mt in to_fetch}
            done = 0
            for fut in concurrent.futures.as_completed(futs):
                done += 1
                dom, sl, mt = futs[fut]
                if progress_cb:
                    progress_cb(f"{item_name}: {done}/{total} site tarandı ({dom})")
                info = fut.result()
                fb = snippet_fallbacks.get(dom)
                if not info.get("status", "").startswith("Aktif"):
                    if fb and fb["emails"]:
                        prim, ordered = rank_emails(fb["emails"], dom)
                        if prim:
                            info = {
                                "company_name": (fb["title"].split("|")[0].split("-")[0].strip() or dom)[:80],
                                "domain": dom, "website": fb["url"] or f"https://{dom}",
                                "primary_email": prim, "all_emails": ordered,
                                "phone": " / ".join(dict.fromkeys(fb["phones"][:2])),
                                "status": "Aktif (Arama Özetinden)", "last_year": "-",
                                "notes": "Arama özetinden kurtarıldı", "product_groups": ", ".join(inferred),
                            }
                        else:
                            continue
                    else:
                        continue
                elif not info.get("primary_email") and fb and fb["emails"]:
                    prim, ordered = rank_emails(fb["emails"], dom)
                    if prim:
                        info["primary_email"], info["all_emails"] = prim, ordered
                if not info.get("primary_email") and not info.get("phone"):
                    continue
                info.update({"item": item_name, "source": sl, "match_type": mt, "last_checked": dt.date.today().isoformat()})
                if info.get("status") == "Aktif":
                    cache[dom] = info
                results.append(info)
    return results


def resolve_mkys_companies(
    company_details: list[dict[str, Any]],
    item_name: str,
    cache: dict[str, dict[str, Any]],
    progress_cb: Any = None,
) -> list[dict[str, Any]]:
    """MKYS dosyasından çıkan firmaları şahıs/kurumsal ayrımıyla güvenli çözer."""
    results: list[dict[str, Any]] = []
    biz_hints = ("ticaret", "medikal", "tedarik", "yapı", "hırdavat", "inşaat", "sanayi", "ltd", "a.ş", "vergi", "satış", "teklif")
    candidates = [c for c in company_details if not c.get("is_public")][:25]
    cache_by_name = {_tr_lower(v.get("company_name", "")): v for v in cache.values() if v.get("company_name")}

    def resolve_one(idx: int, comp: dict[str, Any]) -> dict[str, Any] | None:
        full_name, core, is_person = comp["name"], comp["search_core"], bool(comp.get("is_person"))
        if progress_cb:
            progress_cb(f"MKYS [{idx}/{len(candidates)}]: {core} taranıyor...")
        norm_core = _tr_lower(core)
        for ck, cv in cache_by_name.items():
            if norm_core and norm_core in ck and cv.get("primary_email"):
                hit = dict(cv)
                hit.update({"item": item_name, "source": "MKYS (Önbellek)", "match_type": f"MKYS ({comp['type']})"})
                return hit

        queries = [f'"{full_name}" ticaret medikal tedarik'] if is_person else [f"{core} iletişim", f"{core} medikal"]
        tokens = [t for t in re.split(r"\s+", norm_core) if len(t) >= 3]
        matched_dom, fb_emails, fb_phones, fb_url = "", set(), [], ""
        for q in queries:
            for res in search_multi_engine(q, max_results=5):
                dom = _root_domain(res.get("href", ""))
                if not dom or _is_excluded_domain(dom) or _is_social(dom) or _is_tender_domain(dom):
                    continue
                snip = _tr_lower(f"{res.get('title', '')} {res.get('body', '')}")
                if is_person:
                    if not (tokens and tokens[-1] in snip and tokens[0] in snip and any(b in snip for b in biz_hints)):
                        continue
                elif tokens and tokens[0] not in snip and tokens[0] not in dom:
                    continue
                matched_dom, fb_url = dom, res.get("href", "")
                fb_emails.update(EMAIL_RE.findall(f"{res.get('title', '')} {res.get('body', '')}"))
                fb_phones.extend(PHONE_RE.findall(f"{res.get('title', '')} {res.get('body', '')}"))
                break
            if matched_dom:
                break
        if not matched_dom:
            return None
        info = inspect_company_website(matched_dom, company_hint=full_name, ignore_year_filter=True, page_timeout=3)
        if not info.get("status", "").startswith("Aktif"):
            return None
        if not info.get("primary_email") and fb_emails:
            prim, ordered = rank_emails(fb_emails, matched_dom)
            if prim:
                info["primary_email"], info["all_emails"] = prim, ordered
        if not info.get("primary_email") and not info.get("phone"):
            return None
        info.update({"item": item_name, "company_name": full_name, "source": "MKYS Firma Taraması", "match_type": f"MKYS ({comp['type']})", "last_checked": dt.date.today().isoformat()})
        cache[matched_dom] = info
        return info

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
        for fut in concurrent.futures.as_completed([pool.submit(resolve_one, i, c) for i, c in enumerate(candidates, 1)]):
            item = fut.result()
            if item:
                results.append(item)
    return results


# ---------------------------------------------------------------------------
# Çıktı Yardımcıları (DataFrame, BCC, UTF-8 BOM CSV, Excel)
# ---------------------------------------------------------------------------

def _write_sheet(ws: Any, columns: list[str], rows: list[list[Any]]) -> None:
    fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    ws.append(columns)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = fill, font, Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(max((len(str(c.value or "")) for c in col), default=10) + 3, 14), 55)
    ws.freeze_panes = "A2"


def save_cache(cache_path: Path, cache: dict[str, dict[str, Any]]) -> None:
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Firma Havuzu"
        rows = [
            [
                c["company_name"], c["domain"], c["website"], c["primary_email"],
                "; ".join(e for e in c.get("all_emails", []) if not e.lower().endswith(".kep.tr")),
                c["phone"], c["status"], c.get("last_year", "-"),
                c.get("last_checked", dt.date.today().isoformat()), c.get("notes", ""), c.get("product_groups", ""),
            ]
            for c in sorted(cache.values(), key=lambda x: x["domain"])
        ]
        _write_sheet(ws, CACHE_COLUMNS, rows)
        wb.save(cache_path)
    except Exception:
        pass


def collect_bcc_emails(rows: list[dict[str, Any]]) -> list[str]:
    emails: list[str] = []
    seen: set[str] = set()
    for r in rows:
        prim = (r.get("primary_email") or "").strip().lower()
        cands = ([prim] if prim else []) + [str(e).strip().lower() for e in r.get("all_emails", []) if str(e).strip()]
        for em in cands:
            if not em or em == "-" or em.endswith(".kep.tr") or em in seen:
                continue
            seen.add(em)
            emails.append(em)
    return emails


def rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Kalem": r.get("item", ""),
            "Firma Adı": r.get("company_name", ""),
            "Eşleşme Tipi": r.get("match_type", ""),
            "Öncelikli E-posta": r.get("primary_email", ""),
            "Tüm E-postalar": "; ".join(e for e in r.get("all_emails", []) if not str(e).lower().endswith(".kep.tr")),
            "Telefon": r.get("phone", ""),
            "Web Sitesi": r.get("website", ""),
            "Kaynak": r.get("source", ""),
            "Durum": r.get("status", ""),
        }
        for r in rows
    ])


def dataframe_to_utf8_bom_csv(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)
    return buf.getvalue().encode("utf-8-sig")


def build_excel_bytes(rows: list[dict[str, Any]]) -> bytes:
    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "Tedarikçiler"
    out_rows = [
        [
            r.get("item", ""), r.get("company_name", ""), r.get("match_type", ""), r.get("source", ""),
            r.get("product_groups", ""), r.get("website", ""), r.get("primary_email", ""),
            "; ".join(e for e in r.get("all_emails", []) if not str(e).lower().endswith(".kep.tr")),
            r.get("phone", ""), r.get("status", ""), r.get("last_year", "-"), r.get("notes", ""),
        ]
        for r in rows
    ]
    _write_sheet(ws_main, OUTPUT_COLUMNS, out_rows)
    ws_mails = wb.create_sheet(title="Mail Listesi (Alt Alta)")
    mail_rows = []
    seen_m: set[str] = set()
    for r in rows:
        prim = (r.get("primary_email") or "").strip().lower()
        for idx_em, em in enumerate(([prim] if prim else []) + [str(e).strip().lower() for e in r.get("all_emails", [])]):
            if em and not em.endswith(".kep.tr") and em not in seen_m:
                seen_m.add(em)
                mail_rows.append([em, "Öncelikli" if em == prim else "Alternatif", r.get("item", ""), r.get("company_name", "")])
    _write_sheet(ws_mails, ["E-posta (Kime/BCC'ye Yapıştır)", "Öncelik", "Kalem", "Firma Adı"], mail_rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_rfq_draft(items: list[str]) -> tuple[str, str]:
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


# ---------------------------------------------------------------------------
# Ana Uygulama Akışı
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title="Kamu Satın Alma | Akıllı Tedarikçi & Teklif Havuzu", page_icon="🏥", layout="wide")
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

    st.markdown('<div class="erp-card">', unsafe_allow_html=True)
    col_in, col_file = st.columns([1.45, 1])
    with col_in:
        raw_input = st.text_area(
            "Talep Kalemleri (Her satıra bir kalem yazın veya Excel/Word tablosunu yapıştırın)",
            height=115,
            placeholder="Örn:\nCHİLLER GRUBU İÇİN ANTİFRİZ(ETİLEN GLİKOL)\nDİJİTAL BEBEK BASKÜLÜ",
        )
    with col_file:
        uploaded = st.file_uploader("Talep Formu veya MKYS Dosyası Yükle (.docx, .xlsx, .csv, .pdf)", type=["docx", "xlsx", "xlsm", "csv", "pdf", "txt"])
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
            parsed_file = parse_uploaded_bytes(uploaded.name, uploaded.getvalue())
            if parsed_file.get("error"):
                st.error(parsed_file["error"])
                return
            file_items = [x["item"] for x in parsed_file.get("items", []) if x.get("item")]
            if not items:
                items = file_items or (parsed_file.get("raw_lines") or [])[:15]

        cache = load_cache(CACHE_XLSX)
        ai_cache = load_ai_cache()
        all_rows: list[dict[str, Any]] = []
        intel_map: dict[str, dict[str, Any]] = {}
        progress_bar = st.progress(0)
        status_slot = st.empty()

        if mode_mkys and parsed_file.get("company_details"):
            details = [c for c in parsed_file["company_details"] if not c.get("is_public")]
            item_label = items[0] if items else "MKYS Havuz"
            status_slot.info(f"⏳ MKYS dosyasından {len(details)} firma taranıyor...")
            all_rows = resolve_mkys_companies(details, item_label, cache, progress_cb=lambda m: status_slot.info(f"⏳ {m}"))
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
                item_rows = run_three_ring_search(
                    item, intel, cache, progress_cb=lambda m, i=idx: status_slot.info(f"⏳ [{i}/{total}] {m}"), max_sites=20
                )
                all_rows.extend(item_rows)
                progress_bar.progress(int((idx / total) * 100))

        all_rows.sort(key=lambda r: (r.get("item", ""), 0 if r.get("primary_email") else 1, r.get("company_name", "")))
        save_cache(CACHE_XLSX, cache)
        elapsed = round(time.time() - t0, 1)
        status_slot.success(f"✅ Tarama {elapsed} saniyede tamamlandı ({len(all_rows)} aktif kayıt).")

        st.session_state["last_rows"] = all_rows
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

    k1, k2, k3, k4 = st.columns(4)
    for col, label, val in [
        (k1, "Aranan Kalem", len(last_items)),
        (k2, "Bulunan Firma", len(df)),
        (k3, "Aktif E-Posta", len(bcc_emails)),
        (k4, "İşlem Süresi", f"{elapsed_sec} sn"),
    ]:
        col.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div></div>', unsafe_allow_html=True)

    if intel_map:
        with st.expander("🧠 Yapay Zeka Ürün ve Piyasa Analizleri", expanded=False):
            for item_name, intel in intel_map.items():
                syns = "".join(f'<span class="pill">{s}</span>' for s in intel.get("es_anlamlilar", []))
                brands = "".join(f'<span class="pill">🏷️ {b}</span>' for b in intel.get("markalar", []))
                st.markdown(
                    f'<div class="erp-card"><strong>{item_name}</strong> '
                    f'<span style="color:#0F766E; font-size:0.82rem;">({intel.get("ust_kategori", "")} · {intel.get("source", "")})</span>'
                    f'<p style="margin:6px 0; color:#334155; font-size:0.9rem;">{intel.get("aciklama", "")}</p>'
                    f'<div>{syns} {brands}</div></div>',
                    unsafe_allow_html=True,
                )

    tab1, tab2, tab3 = st.tabs(["🏢 Tedarikçi Tablosu", "✉️ Toplu Mail (BCC) & 4734/22-d Taslağı", "📥 İndir (CSV & Excel)"])

    with tab1:
        if df.empty:
            st.info("Aktif tedarikçi bulunamadı.")
        else:
            st.dataframe(df, use_container_width=True, height=420, hide_index=True, column_config={"Web Sitesi": st.column_config.LinkColumn("Web Sitesi")})

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
                st.download_button("📥 UTF-8 BOM CSV İndir (Türkçe Excel Uyumlu)", data=dataframe_to_utf8_bom_csv(df), file_name="tedarikci_listesi.csv", mime="text/csv", use_container_width=True)
            with c_dl2:
                st.download_button("📊 Excel İndir (tedarikci_listesi.xlsx)", data=build_excel_bytes(rows), file_name="tedarikci_listesi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


if __name__ == "__main__":
    main()

"""Çoklu Arama Motoru (Yandex, Bing, Yahoo, DuckDuckGo, ilan.gov.tr, ihalebul.com) ve Firma Havuzu Motoru.
120 / 240 / 360 site döngüsü altyapısını ve yerel önbellekleri (firma_havuzu.xlsx, ai_cache.json) korur.
"""

from __future__ import annotations

import concurrent.futures
import csv
import datetime as dt
import io
import json
import os
import re
import threading
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urljoin, urlparse

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

# 1. Ziyaret Öncesi Domain Kara Listesi (Pazaryeri, Haber/Medya, İlan/Emlak/Oto/Gezi/Sözlük/Blog)
NEWS_MEDIA_DOMAINS = {
    "hurriyet.com.tr", "hurriyet.com", "milliyet.com.tr", "milliyet.com",
    "sabah.com.tr", "sabah.com", "sozcu.com.tr", "sozcu.com",
    "haberturk.com", "ensonhaber.com", "mynet.com", "onedio.com", "trthaber.com",
    "aa.com.tr", "ntv.com.tr", "cnnturk.com", "cumhuriyet.com.tr", "yenisafak.com",
    "star.com.tr", "aksam.com.tr", "takvim.com.tr", "posta.com.tr", "t24.com.tr",
    "gazeteduvar.com.tr", "bloomberght.com", "dunya.com", "bbc.com", "dw.com",
    "euronews.com", "haberler.com", "sondakika.com", "haber7.com", "internethaber.com",
    "odatv.com", "tele1.com.tr", "birgun.net", "halktv.com.tr", "karar.com",
    "turkiyegazetesi.com.tr", "yenicaggazetesi.com.tr", "fotomac.com.tr", "fanatik.com.tr",
    "sporx.com", "medimagazin.com.tr", "medikalnews.com",
}
CLASSIFIED_BLOG_DOMAINS = {
    "sahibinden.com", "arabam.com", "otosor.com", "hepsiemlak.com", "emlakjet.com",
    "zingat.com", "gezirehberi.com", "tripadvisor.com", "tripadvisor.com.tr",
    "blogspot.com", "wordpress.com", "medium.com", "eksisozluk.com", "uludagsozluk.com",
    "instela.com", "kizlarsoruyor.com", "donanimhaber.com", "technopat.net", "shiftdelete.net",
    "webtekno.com", "wixsite.com", "substack.com", "tumblr.com", "weebly.com",
    "letgo.com", "dolap.com", "modacruz.com", "obilet.com", "enuygun.com",
    "tatilsepeti.com", "etstur.com", "booking.com", "airbnb.com", "trivago.com.tr",
    "biletix.com",
}
MARKETPLACE_DOMAINS = {
    "trendyol.com", "trendyolmilla.com", "hepsiburada.com", "n11.com",
    "amazon.com.tr", "amazon.com", "ciceksepeti.com", "pazarama.com",
    "pttavm.com", "epttavm.com", "gittigidiyor.com", "akakce.com", "cimri.com",
    "idefix.com", "alibaba.com", "aliexpress.com", "ebay.com", "etsy.com",
    "teknosa.com", "vatanbilgisayar.com", "mediamarkt.com.tr", "migros.com.tr",
    "a101.com.tr", "bim.com.tr", "sokmarket.com.tr", "koctas.com.tr",
    "tekzen.com.tr", "tekzen.com", "bauhaus.com.tr", "decathlon.com.tr",
    "morhipo.com", "boyner.com.tr", "kitapyurdu.com", "dr.com.tr", "wikipedia.org",
    "medikalfirma.com", "beymen.com", "ebebek.com", "doktortakvimi.com", "sikayetvar.com",
    "yandex.com.tr", "yandex.com", "google.com", "bing.com", "yahoo.com", "duckduckgo.com",
    "zhihu.com", "researchgate.net", "quora.com", "reddit.com",
    "support.google.com", "apple.com", "microsoft.com",
} | NEWS_MEDIA_DOMAINS | CLASSIFIED_BLOG_DOMAINS

SOCIAL_DOMAINS = {
    "linkedin.com", "instagram.com", "facebook.com", "x.com", "twitter.com",
    "tiktok.com", "youtube.com", "youtu.be", "pinterest.com", "vimeo.com",
    "dailymotion.com", "twitch.tv", "t.me", "telegram.me", "whatsapp.com",
}
DOMAIN_BLOCKED_SUBSTRINGS = ("haber", "gazete", "dergi")
ALLOWED_COMMERCIAL_TLDS = (".tr", ".com", ".net", ".org")

DIRECTORY_DOMAINS = {"firmasec.com", "bulurum.com"}
TENDER_RESULT_DOMAINS = {"ilan.gov.tr", "ihalebul.com"}
TENDER_PORTAL_HINTS = ("ekap.kik.gov.tr", "kik.gov.tr", "ilan.gov.tr", "ihalebul.com", "ihale.com", ".gov.tr", ".bel.tr", ".edu.tr")

DEFAULT_TIMEOUT = 6
DEFAULT_MAX_SITES = 120

MAINTENANCE_KEYWORDS = (
    "bakımda", "bakimda", "yapım aşamasında", "yapim asamasinda",
    "under construction", "coming soon", "alan adı satılık", "alan adı park",
    "park edilmiştir", "park edilmistir", "domain is for sale",
    "this domain may be for sale", "account suspended", "site askıya alındı",
)

# 3. Spam, Bahis ve Kredi Kalıpları
SPAM_BETTING_RE = re.compile(
    r"(?:\$1000|\blening\b|\bcasino\b|\bbahis\b|\bbetting\b|\bslot\b|\bjackpot\b|\bfreespin\b|"
    r"\bdeneme\s+bonusu\b|\bporn\b|\bxxx\b|\bviagra\b|\bcialis\b|\bcrypto\b|\bforex\b|"
    r"\bloan\b|\bpayday\b|\bcanlı\s+bahis\b|\bkaçak\s+bahis\b)",
    re.IGNORECASE,
)
TR_CORP_PAGE_TERMS = (
    "iletişim", "iletisim", "hakkımızda", "hakkimizda", "adres", "tel",
    "vergi", "ticaret", "sanayi", "ltd", "şti", "sti", "a.ş", "a.s",
    "medikal", "kurumsal", "ürünlerimiz", "urunlerimiz",
)
TR_CORP_TITLE_RE = re.compile(
    r"\b(?:limited\s+şirketi|ltd\.?\s*şti\.?|anonim\s+şirketi|a\.?\s*ş\.?|san\.?\s*(?:ve\s*)?tic\.?)\b",
    re.IGNORECASE,
)
TR_PHONE_VALID_RE = re.compile(r"(?:\+?90|0)\s*\(?\s*[235]\d{2}\s*\)?[\s.\-]*\d{3}[\s.\-]*\d{2}[\s.\-]*\d{2}")

EMAIL_RE = re.compile(r"\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,10}\b")
PHONE_RE = re.compile(r"(?:\+90|0)\s*(?:\(?\d{3}\)?[\s.\-]*\d{3}[\s.\-]*\d{2}[\s.\-]*\d{2})")
YEAR_RE = re.compile(r"\b(20[1-3]\d)\b")
COMPANY_NAME_RE = re.compile(
    r"([A-ZÇĞİÖŞÜ][A-Za-zçğıöşüÇĞİÖŞÜ0-9&.,\-\s]{2,95}?"
    r"(?:Limited\s+Şirketi|Ltd\.?\s*Şti\.?|Anonim\s+Şirketi|A\.Ş\.?|San\.?\s*(?:ve\s*)?Tic\.?(?:\s*Ltd\.?\s*Şti\.?)?|Medikal|Tıbbi|Sağlık))"
)
TENDER_CONTRACTOR_LABEL_RE = re.compile(
    r"(?:Yüklenici(?:\s+Adı|\s+Firma|\s+Ünvanı|\s+Unvanı)?|İhaleyi\s+Alan|Kazanan\s+Firma|Sözleşme\s+İmzalanan(?:\s+Firma)?)\s*[:\-]\s*([^\n\r|;]{4,140})",
    re.IGNORECASE,
)
COPYRIGHT_COMPANY_RE = re.compile(
    r"(?:©|&copy;|copyright|tüm\s+hakları\s+saklıdır)\s*(?:20\d{2}(?:\s*[-–]\s*20\d{2})?)?\s*([^\n\r|<]{3,65})",
    re.IGNORECASE,
)

# 4. Genel Sayfa Başlıkları Kara Listesi
GENERIC_PAGE_TITLES = {
    "ana sayfa", "anasayfa", "home", "homepage", "index", "iletisim", "iletişim",
    "contact", "contact us", "bize ulasin", "bize ulaşın", "hakkimizda", "hakkımızda",
    "about", "about us", "urunlerimiz", "ürünlerimiz", "urunler", "ürünler", "products",
    "kategoriler", "categories", "magaza", "mağaza", "shop", "store", "sepetim", "sepet",
    "cart", "404", "sayfa bulunamadi", "sayfa bulunamadı", "not found", "hos geldiniz",
    "hoş geldiniz", "welcome", "giris", "giriş", "login", "blog", "haberler", "kurumsal",
    "hizmetlerimiz", "referanslar", "galeri", "kvkk", "gizlilik politikasi",
}

DOMAIN_SPLIT_WORDS = (
    "medikal", "medical", "saglik", "health", "kimya", "kimyasal", "ticaret", "sanayi",
    "makina", "makine", "teknik", "teknoloji", "grup", "group", "elektronik", "elektrik",
    "mobilya", "dental", "ilac", "ecza", "tekstil", "ambalaj", "gida", "yapi", "hirdavat",
    "otomotiv", "muhendislik", "sistem", "bilisim", "yazilim", "lojistik", "insaat",
    "enerji", "temizlik", "hijyen", "laboratuvar", "tibbi", "cihaz", "sarf", "tedarik",
    "pazarlama", "ithalat", "ihracat", "buro", "ofis", "depo", "deposu", "holding",
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
        "queries": ["yoğun bakım cihazları medikal tedarik"],
    },
    "Ameliyathane & Cerrahi": {
        "keywords": ("laparoskop", "trokar", "bistüri", "koter", "cerrahi", "ameliyat", "sütür", "forseps", "endoskop"),
        "queries": ["cerrahi el aletleri ameliyathane medikal toptan satış"],
    },
    "Kardiyoloji & EKG": {
        "keywords": ("ekg", "elektrokardiyografi", "holter", "efor", "stent", "anjiyo", "kardiyoloji", "tansiyon"),
        "queries": ["kardiyoloji cihazları ekg medikal fiyat teklifi"],
    },
    "Laboratuvar & Patoloji": {
        "keywords": ("santrifüj", "mikroskop", "otoanalizör", "pcr", "pipet", "reaktif", "laboratuvar", "hematoloji", "biyokimya"),
        "queries": ["tıbbi laboratuvar cihazları medikal tedarik"],
    },
    "Görüntüleme & Radyoloji": {
        "keywords": ("\\bultrason\\b", "usg", "röntgen", "tomografi", "emg", "radyoloji", "ultrason prob", "mamografi"),
        "queries": ["tıbbi görüntüleme ultrason röntgen tedarik"],
    },
    "Sterilizasyon & Dezenfeksiyon": {
        "keywords": ("otoklav", "sterilizasyon", "dezenfektan", "plazma", "indikatör", "steril"),
        "queries": ["sterilizasyon otoklav cihazları toptan satış"],
    },
    "Medikal Sarf & Ecza": {
        "keywords": ("enjektör", "eldiven", "kateter", "kanül", "maske", "gazlı bez", "serum", "sonda", "sarf", "ilaç", "ecza", "ampul", "flakon"),
        "queries": ["medikal sarf malzemeleri toptan satış", "ecza deposu iletişim"],
    },
    "Dental / Diş": {
        "keywords": ("\\bdiş\\b", "dental", "diş üniti", "\\bünit\\b", "türbin", "aerotor", "endodonti", "implant"),
        "queries": ["dental cihaz malzeme deposu tedarik"],
    },
    "Biyomedikal Bakım & Kalibrasyon": {
        "keywords": ("bakım", "onarım", "kalibrasyon", "teknik servis", "periyodik", "tıbbi gaz", "biyomedikal"),
        "queries": ["biyomedikal teknik servis kalibrasyon fiyat teklifi"],
    },
    "Fizik Tedavi & Hastane Demirbaş": {
        "keywords": ("fizik tedavi", "ortopedi", "tekerlekli sandalye", "hasta yatağı", "sedye", "muayene masası", "alt değiştirme", "bakım ünitesi", "hastane mobilyası", "tıbbi mobilya"),
        "queries": ["hastane mobilyaları tıbbi demirbaş toptan satış"],
    },
    "Pediatri, Bebek & Ölçüm Cihazları": {
        "keywords": ("bebek", "yenidoğan", "pediatri", "baskül", "terazi", "tartı", "boy ve kilo", "boy kilo", "alt değiştirme", "küvöz"),
        "queries": ["medikal bebek baskülü terazi boy ölçer tedarik"],
    },
    "Endüstriyel Kimya & Glikol": {
        "keywords": ("antifriz", "glikol", "etilen", "monoetilen", "propilen", "kimyasal", "kazan kimyasal", "su şartlandırma", "klor", "asit"),
        "queries": ["etilen glikol monoetilen glikol toptan satış"],
    },
    "İklimlendirme, Chiller & HVAC": {
        "keywords": ("chiller", "klima", "soğutma", "fancoil", "kazan", "brülör", "havalandırma", "kompresör"),
        "queries": ["chiller soğutma kimyasalları antifriz tedarik"],
    },
    "Yapı, Hırdavat & İş Güvenliği": {
        "keywords": ("kaydırmaz", "bant", "boya", "kilit", "hırdavat", "el aleti", "iş güvenliği", "vida", "civata", "matkap"),
        "queries": ['"Yapı ve Hırdavat Malzemeleri" toptan satış "Ltd. Şti." OR "A.Ş." iletişim'],
    },
    "Sıhhi Tesisat & Yapı Malzemeleri": {
        "keywords": ("lavabo", "batarya", "musluk", "klozet", "sifon", "rezervuar", "pvc boru", "vana", "armatür", "tesisat", "vitrifiye", "seramik"),
        "queries": ['"Sıhhi Tesisat ve Yapı Malzemeleri" toptan satış "Ltd. Şti." OR "A.Ş." iletişim'],
    },
    "Büro Makineleri, Kırtasiye & Toner": {
        "keywords": ("toner", "kartuş", "fotokopi", "a4 kağıt", "kağıt", "klasör", "kırtasiye", "kalem", "dosya", "zarf", "baskı"),
        "queries": ['"Büro Makineleri ve Sarf Malzemeleri" toptan satış "Ltd. Şti." OR "A.Ş." iletişim'],
    },
    "Bilişim, Bilgisayar & Ağ Sistemleri": {
        "keywords": ("bilgisayar", "laptop", "monitör", "yazıcı", "tarayıcı", "sunucu", "switch", "router", "ssd", "ram", "barkod", "ups"),
        "queries": ['"Bilişim Teknolojileri ve Donanım" kurumsal tedarik teklif e-posta'],
    },
    "Elektrik & Aydınlatma Malzemeleri": {
        "keywords": ("kablo", "led", "armatür", "priz", "anahtar", "sigorta", "pano", "aydınlatma", "jeneratör", "trafo"),
        "queries": ['"Elektrik ve Aydınlatma Malzemeleri" toptan satış "Ltd. Şti." OR "A.Ş." iletişim'],
    },
    "Ofis & Kurum Mobilyası": {
        "keywords": ("çalışma masası", "ofis koltuğu", "dolap", "keson", "sandalye", "etajer", "bekleme koltuğu", "mobilya"),
        "queries": ['"Ofis ve Kurum Mobilyaları" toptan satış "Ltd. Şti." OR "A.Ş." iletişim'],
    },
    "Temizlik, Hijyen & Gıda": {
        "keywords": ("deterjan", "kağıt havlu", "çöp poşeti", "temizlik", "mutfak", "hijyen", "gıda", "erzak", "çay", "kahve"),
        "queries": ['"Endüstriyel Temizlik ve Hijyen Malzemeleri" toptan satış iletişim'],
    },
}

SERVICE_KEYWORDS = (
    "bakım", "bakim", "onarım", "onarim", "kalibrasyon", "servis", "kiralama",
    "hizmet", "danışmanlık", "egitim", "eğitim", "temizlik", "güvenlik",
    "nakliye", "taşıma", "montaj", "periyodik", "ölçüm", "analiz", "yazılım",
)

OUTPUT_COLUMNS = [
    "Kalem", "Firma Adı", "Sektör", "Eşleşme Tipi", "Kaynak", "Ürün Grupları",
    "Web Sitesi", "Öncelikli E-posta", "Tüm E-postalar", "Telefon", "Durum", "Son Güncelleme Yılı", "Notlar",
]
CACHE_COLUMNS = [
    "Firma Adı", "Domain", "Web Sitesi", "Öncelikli E-posta", "Tüm E-postalar",
    "Telefon", "Durum", "Son Güncelleme Yılı", "Son Kontrol", "Notlar", "Ürün Grupları", "Sektör",
]

_THREAD_LOCAL = threading.local()


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


def tr_lower(text: str) -> str:
    return text.replace("İ", "i").replace("I", "ı").lower()


def tr_upper(text: str) -> str:
    return text.replace("i", "İ").replace("ı", "I").upper()


def _fold_tr(text: str) -> str:
    """Türkçe karakterleri ASCII karşılıklarına katlayarak esnek anahtar kelime eşleşmesi sağlar."""
    low = tr_lower(text)
    return low.translate(str.maketrans("çğıöşü", "cgiosu"))


def detect_procurement_type(items: list[str]) -> str:
    joined = tr_lower(" ".join(items))
    return "Hizmet" if any(kw in joined for kw in SERVICE_KEYWORDS) else "Malzeme"


def _classify_supplier(name: str) -> dict[str, Any]:
    raw = re.sub(r"\s+", " ", str(name)).strip()
    is_pub = bool(PUBLIC_ENTITY_RE.search(tr_upper(raw)))
    is_corp = bool(CORP_INDICATOR_RE.search(raw))
    words = [w for w in re.split(r"[\s\-]+", raw) if w]
    is_person = (not is_corp) and (not is_pub) and (2 <= len(words) <= 4)
    stop_tokens = {"SANAYİ", "SANAYI", "TİCARET", "TICARET", "LİMİTED", "LIMITED", "ŞİRKETİ", "LTD", "ŞTİ", "A.Ş", "VE", "İTHALAT", "İHRACAT", "PAZARLAMA", "SAN", "TİC", "ORT"}
    core_words: list[str] = []
    for w in words:
        w_clean = re.sub(r"[^\wçğıöşüÇĞİÖŞÜ]", "", tr_upper(w))
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


def infer_product_groups(text: str) -> list[str]:
    low = text.lower()
    matched: list[str] = []
    for group_name, cfg in PRODUCT_GROUPS.items():
        for kw in cfg["keywords"]:
            if (kw.startswith("\\b") and re.search(kw, low)) or (not kw.startswith("\\b") and kw in low):
                matched.append(group_name)
                break
    return matched


def build_smart_queries(item_name: str, sector_name: str = "") -> dict[str, Any]:
    """Tedarikçi odaklı genel B2B sorgu kalıpları üretir (toptan satış, ihale yüklenicisi, bölgesel tedarik)."""
    raw = str(item_name).strip()
    paren_parts = [p.strip() for p in re.findall(r"\(([^)]+)\)", raw) if len(p.strip()) >= 3]
    outside = re.sub(r"\([^)]*\)", " ", raw)
    outside_clean = re.sub(
        r"\s+", " ", re.sub(r"\b(için|icin|grubu|tipi|marka|model|adet)\b", " ", outside, flags=re.IGNORECASE)
    ).strip()
    base = outside_clean or raw
    alt = paren_parts[0] if paren_parts else base
    if any("etilen" in tr_lower(p) and "glikol" in tr_lower(p) for p in paren_parts):
        alt = "monoetilen glikol"

    sec = (sector_name or ", ".join(infer_product_groups(item_name)) or "Kurumsal Tedarik").split(",")[0].strip()
    b2b_queries = [
        f'"{sec}" toptan satış "Ltd. Şti." OR "A.Ş." iletişim',
        f'"{base}" "yüklenici" OR "ihale sonucu" sözleşme',
        f'"{sec}" kurumsal tedarik teklif e-posta',
        f'"{sec}" Balıkesir OR İzmir OR Bursa toptan iletişim',
        f"{base} toptan satış tedarikçi iletişim",
        f'{alt} fiyat teklifi iletişim "@"',
    ]
    tender_queries = [
        f"site:ilan.gov.tr {base} ihale sonucu yüklenici",
        f"site:ihalebul.com {base} yüklenici",
    ]
    directory_queries = [
        f"site:firmasec.com {base}",
        f"site:bulurum.com {base}",
    ]
    return {
        "queries": list(dict.fromkeys(b2b_queries)),
        "tender_queries": tender_queries,
        "directory_queries": directory_queries,
        "social_queries": [base],
    }


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
    cache_key = re.sub(r"\s+", " ", item_name.strip().lower())
    inferred_sector = ", ".join(infer_product_groups(item_name)) or "Kurumsal Tedarik"
    fallback_smart = build_smart_queries(item_name, inferred_sector)
    if cache_key in ai_cache:
        cached = dict(ai_cache[cache_key])
        if cached.get("schema_v2") or not api_key:
            if not cached.get("b2b_sorgular"):
                cached["b2b_sorgular"] = fallback_smart["queries"]
            cached["source"] = f"{cached.get('source', 'ai_cache')} (Önbellek)"
            return cached

    from ai_parser import generate_sector_and_b2b_queries

    result = generate_sector_and_b2b_queries(item_name, api_key, fallback_sector=inferred_sector)
    if api_key and result.get("schema_v2"):
        ai_cache[cache_key] = result
        save_ai_cache(ai_cache)
    return result


def _root_domain(url_or_domain: str) -> str:
    raw = url_or_domain.strip().lower()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "https://" + raw
    host = urlparse(raw).netloc.split(":")[0]
    return host[4:] if host.startswith("www.") else host


def _domain_matches(domain: str, domain_set: set[str]) -> bool:
    return any(domain == d or domain.endswith("." + d) for d in domain_set)


def _is_allowed_commercial_tld(domain: str) -> bool:
    """Yalnızca .tr, .com, .net, .org uzantılı ticari alan adlarına izin verir."""
    return bool(domain) and any(domain.endswith(tld) for tld in ALLOWED_COMMERCIAL_TLDS)


def _is_excluded_domain(domain: str) -> bool:
    if not domain:
        return True
    if _domain_matches(domain, MARKETPLACE_DOMAINS):
        return True
    if any(sub in domain for sub in DOMAIN_BLOCKED_SUBSTRINGS):
        return True
    if not _is_allowed_commercial_tld(domain):
        return True
    return False


def _is_social(domain: str) -> bool:
    return _domain_matches(domain, SOCIAL_DOMAINS)


def _is_directory_domain(domain: str) -> bool:
    return _domain_matches(domain, DIRECTORY_DOMAINS)


def _is_tender_result_domain(domain: str) -> bool:
    return _domain_matches(domain, TENDER_RESULT_DOMAINS)


def _is_tender_domain(domain: str) -> bool:
    return any(domain == h or domain.endswith(h) or h in domain for h in TENDER_PORTAL_HINTS)


def _is_pre_fetch_blocked(domain: str) -> bool:
    """Ziyaret öncesi kara liste kontrolü: haber, ilan, blog, sosyal medya ve ticari olmayan TLD'ler."""
    return (
        not domain
        or _is_social(domain)
        or _is_excluded_domain(domain)
        or _is_tender_domain(domain)
    )


def _normalize_phone(raw_phone: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[()]", "", raw_phone)).strip()


def _dedupe_phones(phones: list[str]) -> list[str]:
    seen_digits: set[str] = set()
    out: list[str] = []
    for p in phones:
        norm = _normalize_phone(p)
        digits = re.sub(r"\D", "", norm)
        if digits.startswith("90") and len(digits) == 12:
            digits = "0" + digits[2:]
        if len(digits) >= 10 and digits not in seen_digits:
            seen_digits.add(digits)
            out.append(norm)
    return out


def _standardize_sector(raw_sector: str, groups: list[str] | None = None) -> str:
    parts: list[str] = []
    for piece in re.split(r"[,;|/]+", raw_sector or ""):
        p = re.sub(r"\s+", " ", piece).strip()
        if p and p.lower() not in {x.lower() for x in parts}:
            parts.append(p)
    for g in (groups or []):
        gs = str(g).strip()
        if gs and gs.lower() not in {x.lower() for x in parts}:
            parts.append(gs)
    return ", ".join(parts[:4])


def rank_emails(emails: set[str], site_domain: str) -> tuple[str, list[str]]:
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


def fetch_page(url: str, timeout: int = DEFAULT_TIMEOUT) -> tuple[int | None, str, str]:
    try:
        resp = _get_session().get(url, timeout=timeout, verify=False, allow_redirects=True)
        resp.encoding = resp.apparent_encoding or "utf-8"
        return resp.status_code, resp.url, resp.text
    except requests.RequestException:
        return None, url, ""


# ---------------------------------------------------------------------------
# 4. Akıllı Firma Adı Çıkarma ve Temizleme
# ---------------------------------------------------------------------------

def _is_generic_or_noisy_name(cand: str) -> bool:
    if not cand:
        return True
    cleaned = re.sub(r"\s+", " ", cand).strip(" .,-:;|/•©")
    if len(cleaned) < 2 or len(cleaned) > 50:
        return True
    folded = _fold_tr(cleaned)
    if folded in GENERIC_PAGE_TITLES:
        return True
    if any(folded.startswith(g + " ") or folded.endswith(" " + g) for g in ("ana sayfa", "anasayfa", "home", "sepetim", "404")):
        return True
    if re.search(r"(?:₺|\$|€|\b\d+(?:[.,]\d+)?\s*(?:tl|try|usd|eur)\b)", cleaned, re.IGNORECASE):
        return True
    digits = sum(c.isdigit() for c in cleaned)
    if digits >= 5 and digits / max(len(cleaned), 1) > 0.25:
        return True
    if cleaned.count(",") >= 2:
        return True
    return False


def _sanitize_company_name(raw: str) -> str:
    txt = re.sub(r"\s+", " ", str(raw or "")).strip(" .,-:;|/•©")
    txt = re.sub(r"(?:20\d{2}\s*[-–]?\s*)+$", "", txt).strip(" .,-:;|")
    txt = re.sub(r"\b(?:tüm hakları saklıdır|all rights reserved|copyright)\b.*$", "", txt, flags=re.IGNORECASE).strip(" .,-:;|")
    txt = re.sub(r"(?:\+?90|0\d{3})[\d\s.\-]{6,}", "", txt).strip(" .,-:;|")
    txt = re.sub(r"\b\d+(?:[.,]\d+)?\s*(?:tl|try|usd|eur|₺|\$|€)\b", "", txt, flags=re.IGNORECASE).strip(" .,-:;|")
    return txt[:50].strip(" .,-:;|")


def _humanize_domain_name(domain_or_email: str) -> str:
    raw = str(domain_or_email or "").strip().lower()
    if "@" in raw:
        local, dom = raw.split("@", 1)
        target_dom = dom if dom not in FREE_EMAIL_PROVIDERS else local
    else:
        target_dom = _root_domain(raw) or raw
    stem = target_dom
    for sfx in (".com.tr", ".net.tr", ".org.tr", ".gen.tr", ".biz.tr", ".com", ".net", ".org", ".tr"):
        if stem.endswith(sfx):
            stem = stem[: -len(sfx)]
            break
    stem = stem.replace("-", " ").replace("_", " ").replace(".", " ")
    for w in DOMAIN_SPLIT_WORDS:
        stem = re.sub(rf"([a-z0-9])({w})\b", r"\1 \2", stem)
        stem = re.sub(rf"\b({w})([a-z0-9])", r"\1 \2", stem)
    words = [p.capitalize() for p in stem.split() if p]
    return (" ".join(words) or target_dom)[:50]


def _extract_clean_company_name(
    og_site_name: str,
    copyright_cands: list[str],
    title: str,
    final_dom: str,
    primary_email: str,
    company_hint: str = "",
) -> str:
    # 0) Varsa temiz rehber/ihale ipucu
    hint_clean = _sanitize_company_name(company_hint)
    if hint_clean and not _is_generic_or_noisy_name(hint_clean):
        return hint_clean

    # 1) Telif / ünvan metni ("© 2026 ... Ltd. Şti." / "... A.Ş.") veya og:site_name
    for cc in copyright_cands:
        cand = _sanitize_company_name(cc)
        if cand and not _is_generic_or_noisy_name(cand):
            return cand

    og_clean = _sanitize_company_name(og_site_name)
    if og_clean and not _is_generic_or_noisy_name(og_clean) and "." not in og_clean:
        return og_clean

    # 2) <title> etiketini |, -, –, : işaretlerinden bölüp en temiz marka/firma parçasını al
    if title:
        parts = [ _sanitize_company_name(p) for p in re.split(r"\s*[|\-–:]+\s*", title) ]
        valid_parts = [p for p in parts if p and not _is_generic_or_noisy_name(p)]
        if valid_parts:
            # Kurumsal ek içeren veya daha kısa olan marka parçasını seç
            valid_parts.sort(key=lambda p: (0 if TR_CORP_TITLE_RE.search(p) else 1, len(p)))
            return valid_parts[0]

    # 3) Kurumsal e-posta veya web sitesi alan adından okunaklı firma adı üret
    email_dom = primary_email.split("@")[-1] if ("@" in primary_email and primary_email.split("@")[-1] not in FREE_EMAIL_PROVIDERS) else ""
    return _humanize_domain_name(email_dom or final_dom)


# ---------------------------------------------------------------------------
# 5. Kalem ve Sektör Uyumluluğu Yardımcıları
# ---------------------------------------------------------------------------

def _expand_tr_token_stems(clean_tok: str) -> set[str]:
    out = {clean_tok}
    for sfx in ("lari", "leri", "lar", "ler", "nin", "nun", "si", "su", "i", "u", "e", "a"):
        if clean_tok.endswith(sfx) and len(clean_tok) - len(sfx) >= 4:
            out.add(clean_tok[: -len(sfx)])
            break
    return out


def build_item_relevance_profile(item_name: str, intel: dict[str, Any] | None = None) -> dict[str, Any]:
    """Aranan kalem ve yapay zeka/sektör eş anlamlılarından kök-duyarlı sayfa doğrulama anahtar kelimeleri üretir."""
    intel = intel or {}
    stop = {
        "ve", "veya", "ile", "icin", "grubu", "tipi", "adet", "kutu", "set", "paket",
        "malzeme", "malzemesi", "malzemeleri", "cihaz", "cihazi", "urun", "urunu", "satis", "toptan",
        "tedarik", "tedarikci", "firma", "hizmet", "hizmeti", "alimi", "alımı",
        "adli", "ozel", "genel", "tip", "yeni", "model", "marka",
    }
    raw_terms = [item_name] + list(intel.get("es_anlamlilar") or []) + list(intel.get("markalar") or [])
    core_tokens: set[str] = set()
    for term in raw_terms:
        for w in re.split(r"[\s,/()\-]+", str(term or "")):
            f = re.sub(r"[^a-z0-9]", "", _fold_tr(w))
            if len(f) >= 3 and f not in stop:
                core_tokens.update(_expand_tr_token_stems(f))

    inferred = infer_product_groups(f"{item_name} {intel.get('ust_kategori', '')}")
    sector_tokens: set[str] = set()
    for g in inferred:
        for kw in PRODUCT_GROUPS[g]["keywords"]:
            clean_kw = re.sub(r"\\b", "", kw)
            f_kw = _fold_tr(clean_kw).strip()
            if len(f_kw) >= 3:
                sector_tokens.add(f_kw)
    for w in re.split(r"[\s,/()\-&]+", str(intel.get("ust_kategori") or "")):
        f = re.sub(r"[^a-z0-9]", "", _fold_tr(w))
        if len(f) >= 4 and f not in stop:
            sector_tokens.update(_expand_tr_token_stems(f))

    return {
        "core_tokens": sorted(core_tokens),
        "sector_tokens": sorted(sector_tokens),
        "groups": inferred,
    }


def _is_page_relevant(relevance_text: str, profile: dict[str, Any] | None) -> bool:
    if not profile:
        return True
    core = profile.get("core_tokens") or []
    sector = profile.get("sector_tokens") or []
    if not core and not sector:
        return True
    folded = _fold_tr(relevance_text)
    if any(tok in folded for tok in core):
        return True
    sector_hits = sum(1 for kw in sector if kw in folded)
    return sector_hits >= 2


def _extract_contacts(html: str, base_url: str) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")
    emails: set[str] = set()
    contact_links: list[str] = []
    external_links: list[str] = []
    base_dom = _root_domain(base_url)

    og_el = soup.find("meta", attrs={"property": "og:site_name"}) or soup.find("meta", attrs={"name": "application-name"})
    og_site_name = (og_el.get("content") or "").strip() if og_el else ""
    desc_el = soup.find("meta", attrs={"name": "description"}) or soup.find("meta", attrs={"property": "og:description"})
    meta_desc = (desc_el.get("content") or "").strip() if desc_el else ""

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
        else:
            full = urljoin(base_url, href)
            if urlparse(full).scheme in ("http", "https"):
                ldom = _root_domain(full)
                if ldom and ldom != base_dom and not _is_pre_fetch_blocked(ldom) and not _is_directory_domain(ldom):
                    external_links.append(full)

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = soup.get_text(" ", strip=True)
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    h1_el = soup.find("h1")
    h1_text = h1_el.get_text(" ", strip=True) if h1_el else ""

    # Yan menü / reklam gürültüsü hariç ana içerik metni (sektör uyumluluğu için)
    main_soup = BeautifulSoup(html, "html.parser")
    for tag in main_soup(["script", "style", "noscript", "aside", "nav"]):
        tag.decompose()
    main_el = main_soup.find("main") or main_soup.find("article") or main_soup.body
    main_text = main_el.get_text(" ", strip=True) if main_el else text

    # Footer / telif metninden ünvan adayları
    copyright_cands: list[str] = []
    footer_el = soup.find("footer")
    footer_txt = footer_el.get_text(" ", strip=True) if footer_el else text[-2500:]
    for m in COPYRIGHT_COMPANY_RE.finditer(footer_txt):
        copyright_cands.append(m.group(1))
    for m in COMPANY_NAME_RE.finditer(footer_txt):
        copyright_cands.append(m.group(1))

    emails.update(m.lower() for m in EMAIL_RE.findall(text))
    phones = _dedupe_phones(PHONE_RE.findall(text))
    current_year = dt.date.today().year
    years = [int(y) for y in YEAR_RE.findall(text) if 2010 <= int(y) <= current_year + 1]
    is_maint = any(k in (title + " " + text[:4000]).lower() for k in MAINTENANCE_KEYWORDS)

    return {
        "title": title,
        "h1": h1_text,
        "og_site_name": og_site_name,
        "meta_desc": meta_desc,
        "copyright_cands": copyright_cands,
        "text": text,
        "main_text": main_text,
        "emails": emails,
        "phones": phones,
        "max_year": max(years) if years else None,
        "is_maintenance": is_maint,
        "contact_links": list(dict.fromkeys(contact_links))[:6],
        "external_links": list(dict.fromkeys(external_links))[:5],
        "groups": infer_product_groups(title + " " + main_text[:10000]),
    }


def _has_turkish_corporate_signal(text: str, phones: list[str], final_dom: str) -> bool:
    """Türkçe dil varlığı + kurumsal ifade veya Türkiye telefon numarası doğrulaması."""
    low = tr_lower(text[:30000])
    has_tr_chars = bool(re.search(r"[çğıöşü]", low))
    corp_hits = sum(1 for term in TR_CORP_PAGE_TERMS if term in low)
    has_tr_phone = bool(TR_PHONE_VALID_RE.search(text[:30000])) or any(
        re.sub(r"\D", "", p).startswith(("902", "903", "905", "02", "03", "05")) for p in phones
    )
    if not (has_tr_chars or final_dom.endswith(".tr") or corp_hits >= 2):
        return False
    return corp_hits >= 1 or has_tr_phone


def inspect_company_website(
    dom: str,
    company_hint: str = "",
    ignore_year_filter: bool = False,
    page_timeout: int = DEFAULT_TIMEOUT,
    extra_emails: set[str] | None = None,
    extra_phones: list[str] | None = None,
    relevance_profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if _is_pre_fetch_blocked(dom):
        return {"status": "Elendi (Kara Liste)", "rejected": True}

    status, final_url, html = fetch_page(f"https://{dom}", timeout=page_timeout)
    if status is None:
        status, final_url, html = fetch_page(f"http://{dom}", timeout=page_timeout)
    if status is None or status >= 400 or not html.strip():
        return {"status": "Pasif", "rejected": False}

    final_dom = _root_domain(final_url)
    if _is_pre_fetch_blocked(final_dom):
        return {"status": "Elendi (Kara Liste)", "rejected": True}

    home = _extract_contacts(html, final_url)
    if home["is_maintenance"]:
        return {"status": "Pasif (Bakımda)", "rejected": True}

    combined_home_text = f"{home['title']} {home['meta_desc']} {home['text'][:25000]}"

    # 3a) Yabancı spam / bahis / kredi kesin elemesi
    if SPAM_BETTING_RE.search(combined_home_text):
        return {"status": "Elendi (Spam/Bahis)", "rejected": True}

    # 3b) Türkçe ve Türkiye ticari faaliyet kontrolü
    if not _has_turkish_corporate_signal(combined_home_text, home["phones"], final_dom):
        return {"status": "Elendi (Yurt Dışı / Kurumsal Değil)", "rejected": True}

    # 5) Kalem ve sektör uyumluluğu kontrolü (yan menü/reklam hariç ana içerik)
    relevance_surface = f"{home['title']} {home['h1']} {home['meta_desc']} {home['main_text'][:18000]}"
    if not _is_page_relevant(relevance_surface, relevance_profile):
        return {"status": "Elendi (Alakasız Sektör)", "rejected": True}

    all_emails = set(home["emails"]) | set(extra_emails or ())
    phones = _dedupe_phones(list(home["phones"]) + list(extra_phones or ()))
    years = [home["max_year"]] if home["max_year"] else []
    groups = set(home["groups"])
    copyright_cands = list(home["copyright_cands"])
    combined_all_text = combined_home_text

    cand_links: list[str] = []
    for fb in ("/iletisim", "/contact", "/iletişim", "/kurumsal", "/bize-ulasin"):
        u = urljoin(final_url, fb)
        if u not in cand_links:
            cand_links.append(u)
    for cl in home["contact_links"]:
        if cl not in cand_links:
            cand_links.append(cl)

    scanned_sub = 0
    for link in cand_links:
        if scanned_sub >= 3:
            break
        if _root_domain(link) != final_dom:
            continue
        scanned_sub += 1
        cs, cu, ch = fetch_page(link, timeout=page_timeout)
        if cs and cs < 400 and ch:
            cd = _extract_contacts(ch, cu)
            if SPAM_BETTING_RE.search(cd["text"][:15000]):
                return {"status": "Elendi (Spam/Bahis)", "rejected": True}
            all_emails.update(cd["emails"])
            phones = _dedupe_phones(phones + cd["phones"])
            if cd["max_year"]:
                years.append(cd["max_year"])
            groups.update(cd["groups"])
            copyright_cands.extend(cd["copyright_cands"])
            combined_all_text += " " + cd["text"][:10000]

    max_year = max(years) if years else None
    if max_year is not None and max_year < 2023 and not ignore_year_filter:
        return {"status": f"Pasif ({max_year})", "rejected": True}

    prim, ordered = rank_emails(all_emails, final_dom)

    # 3c) Yalnızca ücretsiz e-posta (@gmail, @hotmail, @yahoo vb.) varsa Türk şirket ünvanı veya +90/02/03/05 telefon şartı
    has_corp_email = any(em.split("@")[-1] not in FREE_EMAIL_PROVIDERS for em in ordered)
    if ordered and not has_corp_email:
        has_corp_title = bool(TR_CORP_TITLE_RE.search(combined_all_text))
        has_tr_phone = bool(TR_PHONE_VALID_RE.search(combined_all_text)) or any(
            re.sub(r"\D", "", p).startswith(("902", "903", "905", "02", "03", "05")) for p in phones
        )
        if not (has_corp_title or has_tr_phone):
            return {"status": "Elendi (Ücretsiz E-posta Doğrulanamadı)", "rejected": True}

    clean_company = _extract_clean_company_name(
        og_site_name=home["og_site_name"],
        copyright_cands=copyright_cands,
        title=home["title"],
        final_dom=final_dom,
        primary_email=prim,
        company_hint=company_hint,
    )

    return {
        "company_name": clean_company,
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


def _unwrap_bing_url(raw_url: str) -> str:
    if "bing.com/ck/a" not in raw_url:
        return raw_url
    try:
        import base64
        u_vals = urlparse(raw_url).query
        u = parse_qs(u_vals).get("u", [""])[0]
        if len(u) > 2:
            b64 = u[2:]
            return base64.urlsafe_b64decode(b64 + "=" * (-len(b64) % 4)).decode("utf-8", errors="ignore")
    except Exception:
        pass
    return raw_url


def _search_yandex(query: str, max_results: int, timeout: int = DEFAULT_TIMEOUT, page: int = 1) -> list[dict[str, str]]:
    try:
        from ddgs.engines.yandex import Yandex
        res = Yandex(timeout=timeout).search(query, region="tr-tr", safesearch="moderate", timelimit=None, page=page)
        return [
            {"title": getattr(r, "title", ""), "href": getattr(r, "href", ""), "body": getattr(r, "body", ""), "engine": "Yandex"}
            for r in (res or [])[:max_results]
        ]
    except Exception:
        return []


def _search_bing(query: str, max_results: int, timeout: int = DEFAULT_TIMEOUT, page: int = 1) -> list[dict[str, str]]:
    if "site:" in query.lower():
        return []
    clean_q = re.sub(r"\([^)]*\)", " ", query).replace('"', " ")
    clean_q = re.sub(r"\s+OR\s+", " ", clean_q, flags=re.IGNORECASE)
    clean_q = re.sub(r"\s+", " ", clean_q).strip()
    if not clean_q:
        return []
    try:
        import xml.etree.ElementTree as ET
        params = {"q": clean_q, "format": "rss", "mkt": "tr-TR", "setlang": "tr", "count": str(max(max_results, 15))}
        if page > 1:
            params["first"] = str((page - 1) * 15 + 1)
        resp = _get_session().get("https://www.bing.com/search", params=params, timeout=timeout, verify=False)
        if resp.status_code != 200 or "<rss" not in resp.text[:400]:
            return []
        root = ET.fromstring(resp.text)
        out: list[dict[str, str]] = []
        for it in root.findall("./channel/item")[:max_results]:
            href = _unwrap_bing_url((it.findtext("link") or "").strip())
            if href:
                out.append({
                    "title": (it.findtext("title") or "").strip(),
                    "href": href,
                    "body": (it.findtext("description") or "").strip(),
                    "engine": "Bing",
                })
        return out
    except Exception:
        return []


def _search_ddgs_backend(query: str, backend: str, max_results: int, timeout: int = DEFAULT_TIMEOUT, page: int = 1) -> list[dict[str, str]]:
    try:
        from ddgs import DDGS
        res = DDGS(timeout=timeout).text(query, region="tr-tr", max_results=max_results, page=page, backend=backend)
        return [{"title": r.get("title", ""), "href": r.get("href", ""), "body": r.get("body", ""), "engine": backend.capitalize()} for r in (res or [])]
    except Exception:
        if backend == "duckduckgo" and page == 1:
            try:
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", RuntimeWarning)
                    from duckduckgo_search import DDGS as LegacyDDGS
                    res = LegacyDDGS(timeout=timeout).text(query, region="tr-tr", max_results=max_results)
                return [{"title": r.get("title", ""), "href": r.get("href", ""), "body": r.get("body", ""), "engine": "DuckDuckGo"} for r in (res or [])]
            except Exception:
                return []
        return []


def search_multi_engine(query: str, max_results: int = 15, timeout: int = DEFAULT_TIMEOUT, page: int = 1) -> list[dict[str, str]]:
    """Yandex, Bing, Yahoo ve DuckDuckGo arama kaynaklarını paralel sorgular ve birleştirir."""
    merged: list[dict[str, str]] = []
    seen: dict[str, dict[str, str]] = {}
    tasks = [
        (_search_yandex, (query, max_results, timeout, page)),
        (_search_bing, (query, max_results, timeout, page)),
        (_search_ddgs_backend, (query, "yahoo", max_results, timeout, page)),
        (_search_ddgs_backend, (query, "duckduckgo", max_results, timeout, page)),
    ]
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        futs = [pool.submit(fn, *args) for fn, args in tasks]
        for fut in concurrent.futures.as_completed(futs):
            for r in fut.result():
                href = (r.get("href") or "").strip()
                if not href:
                    continue
                if href not in seen:
                    seen[href] = r
                    merged.append(r)
                else:
                    extra_body = (r.get("body") or "").strip()
                    if extra_body and extra_body not in seen[href].get("body", ""):
                        seen[href]["body"] = f"{seen[href].get('body', '')} {extra_body}".strip()
    return merged


def load_cache(cache_path: Path = CACHE_XLSX) -> dict[str, dict[str, Any]]:
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
            pgroups = str(r[10] or "") if len(r) > 10 else ""
            sector = str(r[11] or "") if len(r) > 11 else ""
            comp = str(r[0] or dom)
            if _is_generic_or_noisy_name(comp):
                comp = _humanize_domain_name(prim or dom)
            cache[dom] = {
                "company_name": comp, "domain": dom, "website": str(r[2] or ""),
                "primary_email": prim, "all_emails": all_em,
                "phone": " / ".join(_dedupe_phones(re.split(r"[/,;]+", str(r[5] or "")))[:2]),
                "status": str(r[6] or ""), "last_year": r[7] or "-", "last_checked": str(r[8] or ""),
                "notes": str(r[9] or ""), "product_groups": pgroups,
                "sector": sector or pgroups,
            }
    except Exception:
        pass
    return cache


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
                c.get("last_checked", dt.date.today().isoformat()), c.get("notes", ""),
                c.get("product_groups", ""), c.get("sector", c.get("product_groups", "")),
            ]
            for c in sorted(cache.values(), key=lambda x: x["domain"])
        ]
        _write_sheet(ws, CACHE_COLUMNS, rows)
        wb.save(cache_path)
    except Exception:
        pass


def _merge_firm_into_cache(cache: dict[str, dict[str, Any]], info: dict[str, Any], sector_tag: str, inferred: list[str]) -> None:
    dom = (info.get("domain") or _root_domain(info.get("website", ""))).strip().lower()
    if not dom:
        return
    existing = cache.get(dom, {})
    merged_emails = set(existing.get("all_emails", [])) | set(info.get("all_emails", []))
    if existing.get("primary_email"):
        merged_emails.add(existing["primary_email"])
    if info.get("primary_email"):
        merged_emails.add(info["primary_email"])
    prim, ordered = rank_emails(merged_emails, dom)

    raw_phones = re.split(r"[/,;]+", f"{info.get('phone', '')} / {existing.get('phone', '')}")
    phones = _dedupe_phones(raw_phones)

    merged_sector = _standardize_sector(f"{existing.get('sector', '')}, {info.get('sector', '')}, {sector_tag}", inferred)
    merged_groups = _standardize_sector(f"{existing.get('product_groups', '')}, {info.get('product_groups', '')}", inferred)
    info["primary_email"] = prim or info.get("primary_email", "")
    info["all_emails"] = ordered
    info["phone"] = " / ".join(phones[:2])
    info["sector"] = merged_sector
    info["product_groups"] = merged_groups

    comp_name = info.get("company_name") or existing.get("company_name") or dom
    if _is_generic_or_noisy_name(comp_name):
        comp_name = _humanize_domain_name(info["primary_email"] or dom)

    cache[dom] = {
        "company_name": comp_name,
        "domain": dom,
        "website": info.get("website") or existing.get("website") or f"https://{dom}",
        "primary_email": info["primary_email"],
        "all_emails": ordered,
        "phone": info["phone"],
        "status": info.get("status") or existing.get("status") or "Aktif",
        "last_year": info.get("last_year") or existing.get("last_year") or "-",
        "last_checked": dt.date.today().isoformat(),
        "notes": info.get("notes") or existing.get("notes") or "",
        "product_groups": merged_groups,
        "sector": merged_sector,
    }
    try:
        import db
        db.upsert_master_supplier(cache[dom], sector=sector_tag, item_name=str(info.get("item") or ""))
    except Exception:
        pass


def search_cache_pool(item_name: str, intel: dict[str, Any], cache: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    ust_kat = intel.get("ust_kategori") or ""
    inferred = infer_product_groups(f"{item_name} {ust_kat}")
    stop = {"ve", "ile", "için", "icin", "grubu", "tipi", "adet", "satış", "satis", "toptan", "tedarik", "tedarikçi", "firma", "medikal", "ürün", "cihaz"}
    raw_terms = [item_name, ust_kat] + list(intel.get("es_anlamlilar") or [])
    tokens = {
        re.sub(r"[^\wçğıöşü]", "", tr_lower(w))
        for t in raw_terms
        for w in re.split(r"[\s,/()\-]+", str(t))
    }
    tokens = {t for t in tokens if len(t) >= 3 and t not in stop}

    matched: list[tuple[int, dict[str, Any]]] = []
    for dom, entry in cache.items():
        if _is_pre_fetch_blocked(dom):
            continue
        if not str(entry.get("status", "")).startswith("Aktif"):
            continue
        if not entry.get("primary_email") and not entry.get("phone"):
            continue
        haystack = tr_lower(" ".join([
            entry.get("company_name", ""),
            entry.get("sector", ""),
            entry.get("product_groups", ""),
            entry.get("notes", ""),
        ]))
        group_hits = sum(1 for g in inferred if tr_lower(g) in haystack)
        token_hits = sum(1 for tk in tokens if tk in haystack)
        score = group_hits * 3 + token_hits
        if score <= 0:
            continue
        hit = dict(entry)
        hit.update({
            "item": item_name,
            "source": "Firma Havuzu (Önbellek)",
            "match_type": "Havuz Eşleşmesi (Sektör)",
        })
        matched.append((score, hit))

    matched.sort(key=lambda pair: (-pair[0], 0 if pair[1].get("primary_email") else 1, pair[1].get("company_name", "")))
    return [h for _, h in matched]


def _extract_directory_leads(
    dir_items: list[dict[str, str]],
    item_name: str,
    sector_tag: str,
    inferred: list[str],
    timeout: int = DEFAULT_TIMEOUT,
) -> tuple[list[tuple[str, str, str, set[str], list[str], str]], list[dict[str, Any]]]:
    web_candidates: list[tuple[str, str, str, set[str], list[str], str]] = []
    direct_leads: list[dict[str, Any]] = []
    if not dir_items:
        return web_candidates, direct_leads

    def parse_one(res: dict[str, str]) -> dict[str, Any]:
        href = res.get("href", "")
        title = res.get("title", "")
        body = res.get("body", "")
        snip_txt = f"{title} {body}"
        emails = {m.lower() for m in EMAIL_RE.findall(snip_txt)}
        phones = _dedupe_phones(PHONE_RE.findall(snip_txt))
        comp_name = re.split(r"\s*[|/\-]\s*(?:firmasec|bulurum)", title, flags=re.IGNORECASE)[0].strip()
        ext_doms: list[str] = []
        if href:
            st_code, f_url, html = fetch_page(href, timeout=timeout)
            if st_code and st_code < 400 and html:
                cd = _extract_contacts(html, f_url)
                emails.update(cd["emails"])
                phones = _dedupe_phones(phones + cd["phones"])
                if cd["h1"] and len(cd["h1"]) > 2:
                    comp_name = cd["h1"]
                for ext_u in cd["external_links"]:
                    ed = _root_domain(ext_u)
                    if ed and not _is_pre_fetch_blocked(ed):
                        ext_doms.append(ed)
        cleaned_comp = _sanitize_company_name(comp_name)
        return {
            "href": href,
            "comp_name": cleaned_comp if not _is_generic_or_noisy_name(cleaned_comp) else "Rehber Firması",
            "emails": emails,
            "phones": phones,
            "ext_doms": list(dict.fromkeys(ext_doms)),
        }

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
        for parsed in pool.map(parse_one, dir_items[:15]):
            if parsed["ext_doms"]:
                web_candidates.append((
                    parsed["ext_doms"][0], "Rehber Taraması (firmasec/bulurum)", "Rehber Eşleşmesi",
                    parsed["emails"], parsed["phones"], parsed["comp_name"],
                ))
            else:
                dom = _root_domain(parsed["href"])
                prim, ordered = rank_emails(parsed["emails"], "")
                if prim or parsed["phones"]:
                    direct_leads.append({
                        "item": item_name,
                        "company_name": parsed["comp_name"],
                        "domain": f"{dom}/{re.sub(r'[^a-z0-9]', '', tr_lower(parsed['comp_name']))[:24]}" if dom else parsed["comp_name"],
                        "website": parsed["href"],
                        "primary_email": prim,
                        "all_emails": ordered,
                        "phone": " / ".join(parsed["phones"][:2]),
                        "status": "Aktif (Rehber)",
                        "last_year": "-",
                        "notes": "Rehber sayfasından çıkarıldı",
                        "source": "Rehber Taraması (firmasec/bulurum)",
                        "match_type": "Rehber Eşleşmesi",
                        "sector": sector_tag,
                        "product_groups": ", ".join(inferred),
                    })
    return web_candidates, direct_leads


def _extract_contractor_names_from_text(text: str) -> list[str]:
    found: list[str] = []
    for m in TENDER_CONTRACTOR_LABEL_RE.finditer(text):
        cand = re.sub(r"\s+", " ", m.group(1)).strip(" .,-:;")
        cand = re.split(r"\s+(?:Sözleşme|İhale|Bedel|Tarih|Tutarı|TL|TRY)\b", cand, flags=re.IGNORECASE)[0].strip(" .,-:;")
        if len(cand) >= 5 and not PUBLIC_ENTITY_RE.search(tr_upper(cand)):
            found.append(cand[:110])
    for m in COMPANY_NAME_RE.finditer(text):
        cand = re.sub(r"\s+", " ", m.group(1)).strip(" .,-:;")
        if len(cand) >= 6 and not PUBLIC_ENTITY_RE.search(tr_upper(cand)) and not BOILERPLATE_LINE_RE.match(cand):
            found.append(cand[:110])
    unique: list[str] = []
    seen: set[str] = set()
    for c in found:
        key = tr_lower(c)
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique


def _resolve_tender_contractors(
    tender_items: list[dict[str, str]],
    item_name: str,
    sector_tag: str,
    inferred: list[str],
    cache: dict[str, dict[str, Any]],
    progress_cb: Any = None,
    timeout: int = DEFAULT_TIMEOUT,
) -> list[dict[str, Any]]:
    if not tender_items:
        return []
    contractor_names: list[str] = []

    def scan_tender_item(res: dict[str, str]) -> list[str]:
        txt = f"{res.get('title', '')}\n{res.get('body', '')}"
        href = res.get("href", "")
        if href:
            sc, _, html = fetch_page(href, timeout=timeout)
            if sc and sc < 400 and len(html) > 500:
                soup = BeautifulSoup(html, "html.parser")
                for tag in soup(["script", "style", "noscript"]):
                    tag.decompose()
                page_txt = soup.get_text(" ", strip=True)
                if len(page_txt) > 80:
                    txt = f"{txt}\n{page_txt[:15000]}"
        return _extract_contractor_names_from_text(txt)

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
        for names in pool.map(scan_tender_item, tender_items[:16]):
            contractor_names.extend(names)

    deduped: list[str] = []
    seen_cores: set[str] = set()
    for n in contractor_names:
        cl = _classify_supplier(n)
        if cl.get("is_public"):
            continue
        core_key = tr_lower(cl["search_core"])
        if core_key and core_key not in seen_cores:
            seen_cores.add(core_key)
            deduped.append(n)
        if len(deduped) >= 12:
            break

    if not deduped:
        return []
    if progress_cb:
        progress_cb(f"{item_name}: {len(deduped)} ihale yüklenicisinin iletişim bilgileri aranıyor...")

    cache_by_name = {tr_lower(v.get("company_name", "")): v for v in cache.values() if v.get("company_name")}
    results: list[dict[str, Any]] = []

    def resolve_contractor(comp_name: str) -> dict[str, Any] | None:
        cl = _classify_supplier(comp_name)
        core = cl["search_core"]
        norm_core = tr_lower(core)
        for ck, cv in cache_by_name.items():
            if norm_core and norm_core in ck and (cv.get("primary_email") or cv.get("phone")):
                hit = dict(cv)
                hit.update({
                    "item": item_name,
                    "source": "İhale Sonuç Taraması (Önbellek)",
                    "match_type": "İhale Yüklenicisi (ilan.gov.tr / ihalebul.com)",
                })
                return hit

        tokens = [t for t in re.split(r"\s+", norm_core) if len(t) >= 3]
        matched_dom, fb_emails, fb_phones, fb_url = "", set(), [], ""
        for q in (f'"{core}" iletişim teklif', f"{core} iletişim"):
            for r in search_multi_engine(q, max_results=6, timeout=timeout):
                dom = _root_domain(r.get("href", ""))
                if _is_pre_fetch_blocked(dom) or _is_directory_domain(dom):
                    continue
                snip = tr_lower(f"{r.get('title', '')} {r.get('body', '')}")
                if tokens and tokens[0] not in snip and tokens[0] not in dom:
                    continue
                matched_dom, fb_url = dom, r.get("href", "")
                raw_snip = f"{r.get('title', '')} {r.get('body', '')}"
                fb_emails.update(m.lower() for m in EMAIL_RE.findall(raw_snip))
                fb_phones.extend(_dedupe_phones(PHONE_RE.findall(raw_snip)))
                break
            if matched_dom:
                break

        if not matched_dom:
            return None
        info = inspect_company_website(
            matched_dom, company_hint=comp_name, ignore_year_filter=True,
            page_timeout=timeout, extra_emails=fb_emails, extra_phones=fb_phones,
        )
        if not info.get("status", "").startswith("Aktif"):
            if info.get("rejected"):
                return None
            prim, ordered = rank_emails(fb_emails, matched_dom)
            if prim or fb_phones:
                info = {
                    "company_name": _sanitize_company_name(comp_name) or _humanize_domain_name(matched_dom),
                    "domain": matched_dom,
                    "website": fb_url or f"https://{matched_dom}",
                    "primary_email": prim, "all_emails": ordered,
                    "phone": " / ".join(_dedupe_phones(fb_phones)[:2]),
                    "status": "Aktif (İhale Özetinden)", "last_year": "-",
                    "notes": "İhale yüklenicisi özetinden kurtarıldı",
                }
            else:
                return None
        if not info.get("primary_email") and not info.get("phone"):
            return None
        info.update({
            "item": item_name,
            "company_name": _sanitize_company_name(comp_name) or info.get("company_name") or _humanize_domain_name(matched_dom),
            "source": "İhale Sonuç Taraması",
            "match_type": "İhale Yüklenicisi (ilan.gov.tr / ihalebul.com)",
            "sector": sector_tag,
            "product_groups": _standardize_sector(info.get("product_groups", ""), inferred),
        })
        return info

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        for res in pool.map(resolve_contractor, deduped):
            if res:
                results.append(res)
    return results


def _build_round_queries(
    item_name: str, intel: dict[str, Any], round_num: int
) -> tuple[list[tuple[str, str, str, int]], list[str]]:
    ust_kat = (intel.get("ust_kategori") or "").split(",")[0].strip()
    inferred = infer_product_groups(f"{item_name} {ust_kat}")
    smart = build_smart_queries(item_name, ust_kat)
    syns = [s for s in (intel.get("es_anlamlilar") or []) if s]
    brands = [b for b in (intel.get("markalar") or []) if b]

    if round_num <= 1:
        gemini_b2b = [q for q in (intel.get("b2b_sorgular") or []) if q]
        ring1 = list(dict.fromkeys(gemini_b2b + smart["queries"]))[:6]
        ring2 = [f'"{ust_kat}" toptan tedarik iletişim'] if ust_kat else []
        for g in inferred[:2]:
            ring2.extend(PRODUCT_GROUPS[g]["queries"][:1])
        ring3 = [f"{m} türkiye distribütör yetkili satıcı iletişim" for m in brands[:2]]
        specs = (
            [(q, "Web Arama (B2B)", "1. Halka (B2B Sektör & Toptan)", 15) for q in ring1]
            + [(q, "Web Arama (Kategori)", "2. Halka (Üst Sektör)", 12) for q in list(dict.fromkeys(ring2))[:3]]
            + [(q, "Marka Taraması", "3. Halka (Marka)", 10) for q in ring3]
            + [(q, "Rehber Taraması", "Rehber (firmasec/bulurum)", 12) for q in smart["directory_queries"]]
            + [(q, "İhale Sonuç Taraması", "İhale (ilan.gov.tr/ihalebul)", 10) for q in smart["tender_queries"]]
        )
        return specs, inferred

    syn_a = syns[0] if syns else item_name
    syn_b = syns[1] if len(syns) > 1 else syn_a
    if round_num == 2:
        extra_b2b = [
            f'"{ust_kat or syn_a}" toptan satış "Ltd. Şti." OR "A.Ş." iletişim',
            f'"{syn_a}" "yüklenici" OR "ihale sonucu" sözleşme',
            f'"{ust_kat or syn_b}" kurumsal tedarik teklif e-posta',
            f'"{ust_kat or syn_a}" Balıkesir OR İzmir OR Bursa toptan iletişim',
            f'{syn_b} toptan satış fiyat teklifi iletişim "@"',
            f"{brands[0]} türkiye yetkili satıcı iletişim" if brands else f"{item_name} ithalatçı firma iletişim",
        ]
        extra_dir = [f"site:firmasec.com {syn_a}", f"site:bulurum.com {syn_a}"]
        extra_tender = [f"site:ilan.gov.tr {syn_a} yüklenici", f"site:ihalebul.com {syn_a}"]
    else:
        cities = ["balıkesir", "izmir", "bursa", "istanbul", "ankara", "manisa"]
        city_hint = cities[(round_num - 3) % len(cities)]
        syn_pick = syns[(round_num - 1) % len(syns)] if syns else item_name
        extra_b2b = [
            f'"{ust_kat or syn_pick}" {city_hint} toptan satış iletişim',
            f'"{syn_pick}" toptan satış "Ltd. Şti." OR "A.Ş." {city_hint}',
            f'{syn_pick} kurumsal tedarik teklif "@" {city_hint}',
            f'"{ust_kat or syn_pick}" fiyat teklifi türkiye',
            f"{brands[(round_num - 1) % len(brands)]} türkiye distribütör" if brands else f"{syn_pick} imalatçı firma iletişim",
        ]
        extra_dir = [f"site:firmasec.com {syn_pick} {city_hint}", f"site:bulurum.com {syn_pick} {city_hint}"]
        extra_tender = [f"site:ilan.gov.tr {syn_pick}", f"site:ihalebul.com {syn_pick} sonuç"]

    specs = (
        [(q, f"Web Arama ({round_num}. Tur)", f"{round_num}. Tur (B2B Sorgu)", 15) for q in list(dict.fromkeys(extra_b2b))]
        + [(q, f"Rehber ({round_num}. Tur)", "Rehber (firmasec/bulurum)", 12) for q in extra_dir]
        + [(q, f"İhale ({round_num}. Tur)", "İhale (ilan.gov.tr/ihalebul)", 10) for q in extra_tender]
    )
    return specs, inferred


def _resolve_mkys_priority_leads(
    item_name: str,
    sector_tag: str,
    inferred: list[str],
    synonyms: list[str],
    cache: dict[str, dict[str, Any]],
    progress_cb: Any = None,
    timeout: int = DEFAULT_TIMEOUT,
) -> list[dict[str, Any]]:
    """Önce MKYS geçmiş alımlarında ve Ana Firma Havuzunda eşleşen firmaları bulur; iletişimi eksik olanları web'den tamamlar."""
    try:
        import db
        matched = db.query_mkys_and_master_suppliers(item_name, sector=sector_tag, synonyms=synonyms)
    except Exception:
        matched = []
    if not matched:
        return []

    ready: list[dict[str, Any]] = []
    to_lookup: list[dict[str, Any]] = []

    for m in matched[:30]:
        if not m.get("needs_contact_lookup"):
            rec = {
                "item": item_name,
                "company_name": m["company_name"],
                "domain": m.get("domain") or "",
                "website": m.get("website") or "",
                "primary_email": m.get("primary_email") or "",
                "all_emails": m.get("all_emails") or [],
                "phone": m.get("phone") or "",
                "status": "Aktif (Havuz/MKYS)",
                "last_year": "-",
                "notes": f"Geçmiş kalemler: {', '.join((m.get('matched_items') or [])[:3])}",
                "source": m.get("source") or "MKYS Geçmiş Alım",
                "match_type": m.get("match_type") or "MKYS / Havuz Eşleşmesi",
                "sector": sector_tag or m.get("sector") or "Kurumsal Tedarik",
                "product_groups": _standardize_sector(m.get("sector", ""), inferred),
            }
            ready.append(rec)
        else:
            to_lookup.append(m)

    to_lookup = to_lookup[:12]
    if to_lookup and progress_cb:
        progress_cb(f"{item_name}: MKYS/Havuz'dan {len(to_lookup)} geçmiş tedarikçinin web/e-posta bilgileri tamamlanıyor...")

    def lookup_one(m: dict[str, Any]) -> dict[str, Any]:
        comp_name = m["company_name"]
        cl = _classify_supplier(comp_name)
        core = cl["search_core"]
        norm_core = tr_lower(core)
        tokens = [t for t in re.split(r"\s+", norm_core) if len(t) >= 3]
        matched_dom, fb_emails, fb_phones, fb_url = "", set(), [], ""

        for q in (f'"{core}" iletişim teklif', f"{core} iletişim"):
            for r in search_multi_engine(q, max_results=5, timeout=timeout):
                dom = _root_domain(r.get("href", ""))
                if _is_pre_fetch_blocked(dom) or _is_directory_domain(dom):
                    continue
                snip = tr_lower(f"{r.get('title', '')} {r.get('body', '')}")
                if tokens and tokens[0] not in snip and tokens[0] not in dom:
                    continue
                matched_dom, fb_url = dom, r.get("href", "")
                raw_snip = f"{r.get('title', '')} {r.get('body', '')}"
                fb_emails.update(m_em.lower() for m_em in EMAIL_RE.findall(raw_snip))
                fb_phones.extend(_dedupe_phones(PHONE_RE.findall(raw_snip)))
                break
            if matched_dom:
                break

        info: dict[str, Any] = {}
        if matched_dom:
            info = inspect_company_website(
                matched_dom,
                company_hint=comp_name,
                ignore_year_filter=True,
                page_timeout=timeout,
                extra_emails=fb_emails,
                extra_phones=fb_phones,
                relevance_profile=None,
            )
        if not info.get("status", "").startswith("Aktif"):
            prim, ordered = rank_emails(fb_emails, matched_dom)
            info = {
                "company_name": comp_name,
                "domain": matched_dom,
                "website": fb_url or (f"https://{matched_dom}" if matched_dom else ""),
                "primary_email": prim,
                "all_emails": ordered,
                "phone": " / ".join(_dedupe_phones(fb_phones)[:2]),
                "status": "Aktif (MKYS + Web)" if (prim or fb_phones) else "MKYS Kayıtlı (İletişim Bekliyor)",
                "last_year": "-",
                "notes": f"Geçmiş alım: {', '.join((m.get('matched_items') or [])[:2])}",
            }
        info.update({
            "item": item_name,
            "company_name": comp_name,
            "source": "MKYS Geçmiş Alım (Web Tamamlandı)" if (info.get("primary_email") or info.get("phone")) else "MKYS Geçmiş Alım",
            "match_type": m.get("match_type") or "MKYS Geçmiş Alım",
            "sector": sector_tag or m.get("sector") or "Kurumsal Tedarik",
            "product_groups": _standardize_sector(m.get("sector", ""), inferred),
        })
        if info.get("domain"):
            _merge_firm_into_cache(cache, info, info["sector"], inferred)
        else:
            try:
                import db
                db.upsert_master_supplier(info, sector=info["sector"], item_name=item_name)
            except Exception:
                pass
        return info

    if to_lookup:
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
            for res in pool.map(lookup_one, to_lookup):
                ready.append(res)

    return ready


def run_three_ring_search(
    item_name: str,
    intel: dict[str, Any],
    cache: dict[str, dict[str, Any]],
    progress_cb: Any = None,
    max_sites: int = DEFAULT_MAX_SITES,
    round_num: int = 1,
    skip_domains: set[str] | None = None,
) -> list[dict[str, Any]]:
    """Önce MKYS geçmiş alımları ve Ana Firma Havuzunu getirip eksikleri web'de tamamlar; ardından B2B 4 motor + rehber + ihale taraması yapar."""
    ust_kat = (intel.get("ust_kategori") or "").split(",")[0].strip()
    query_specs, inferred = _build_round_queries(item_name, intel, round_num)
    sector_tag = ust_kat or _standardize_sector("", inferred) or "Kurumsal Tedarik"
    relevance_profile = build_item_relevance_profile(item_name, intel)

    seen_domains: set[str] = set(skip_domains or ())
    seen_firms: set[str] = set()
    results: list[dict[str, Any]] = []

    if round_num <= 1:
        # 1) Öncelikli olarak SQLite MKYS geçmiş alımları ve Ana Firma Havuzu (+ eksik iletişim tamamlama)
        mkys_leads = _resolve_mkys_priority_leads(
            item_name=item_name,
            sector_tag=sector_tag,
            inferred=inferred,
            synonyms=list(intel.get("es_anlamlilar") or []),
            cache=cache,
            progress_cb=progress_cb,
            timeout=DEFAULT_TIMEOUT,
        )
        for ml in mkys_leads:
            mdom = str(ml.get("domain", "")).lower()
            mfirm = tr_lower(str(ml.get("company_name", "")))
            if mdom:
                seen_domains.add(mdom)
                if skip_domains is not None:
                    skip_domains.add(mdom)
            if mfirm:
                seen_firms.add(mfirm)
            results.append(ml)

        # 2) firma_havuzu.xlsx içindeki eşleşen firmalar
        pool_hits = search_cache_pool(item_name, intel, cache)
        for ph in pool_hits:
            pdom = str(ph.get("domain", "")).lower()
            pfirm = tr_lower(str(ph.get("company_name", "")))
            if (pdom and pdom in seen_domains) or (pfirm and pfirm in seen_firms):
                continue
            if pdom:
                seen_domains.add(pdom)
            if pfirm:
                seen_firms.add(pfirm)
            ph["sector"] = ph.get("sector") or sector_tag
            results.append(ph)

    if progress_cb:
        progress_cb(f"{item_name} ({round_num}. Tur): {len(results)} MKYS/havuz kaydı · {len(query_specs)} B2B sorgu 4 motorda taranıyor...")

    candidate_domains: dict[str, tuple[str, str]] = {}
    snippet_fallbacks: dict[str, dict[str, Any]] = {}
    directory_hits: list[dict[str, str]] = []
    tender_hits: list[dict[str, str]] = []
    seen_special_urls: set[str] = set()
    search_page = max(1, round_num)

    with concurrent.futures.ThreadPoolExecutor(max_workers=6) as qpool:
        q_futs = {qpool.submit(search_multi_engine, q, mr, DEFAULT_TIMEOUT, search_page): (sl, mt) for q, sl, mt, mr in query_specs}
        for qfut in concurrent.futures.as_completed(q_futs):
            sl, mt = q_futs[qfut]
            for res in qfut.result():
                href = res.get("href", "")
                dom = _root_domain(href)
                if not dom:
                    continue
                if _is_tender_result_domain(dom):
                    if href and href not in seen_special_urls:
                        seen_special_urls.add(href)
                        tender_hits.append(res)
                    continue
                if _is_directory_domain(dom):
                    if href and href not in seen_special_urls:
                        seen_special_urls.add(href)
                        directory_hits.append(res)
                    continue
                # 1) Ziyaret öncesi domain kara listesi (haber, ilan, blog, sosyal medya ve ticari olmayan TLD)
                if _is_pre_fetch_blocked(dom):
                    continue
                if dom in seen_domains:
                    continue

                txt = f"{res.get('title', '')} {res.get('body', '')}"
                if SPAM_BETTING_RE.search(txt):
                    continue

                candidate_domains.setdefault(dom, (sl, mt))
                s_em = {m.lower() for m in EMAIL_RE.findall(txt)}
                s_ph = _dedupe_phones(PHONE_RE.findall(txt))
                fb = snippet_fallbacks.setdefault(dom, {"emails": set(), "phones": [], "title": res.get("title", ""), "url": href})
                fb["emails"].update(s_em)
                fb["phones"] = _dedupe_phones(fb["phones"] + s_ph)

    dir_web_cands, dir_direct_leads = _extract_directory_leads(directory_hits, item_name, sector_tag, inferred, timeout=DEFAULT_TIMEOUT)
    for edom, d_sl, d_mt, d_em, d_ph, d_hint in dir_web_cands:
        if edom in seen_domains or _is_pre_fetch_blocked(edom):
            continue
        candidate_domains.setdefault(edom, (d_sl, d_mt))
        fb = snippet_fallbacks.setdefault(edom, {"emails": set(), "phones": [], "title": d_hint, "url": f"https://{edom}"})
        fb["emails"].update(d_em)
        fb["phones"] = _dedupe_phones(fb["phones"] + d_ph)
        fb["company_hint"] = d_hint

    for lead in dir_direct_leads:
        ldom = lead["domain"]
        if ldom not in seen_domains:
            seen_domains.add(ldom)
            if skip_domains is not None:
                skip_domains.add(ldom)
            _merge_firm_into_cache(cache, lead, sector_tag, inferred)
            results.append(lead)

    to_fetch: list[tuple[str, str, str]] = []
    for dom, (sl, mt) in candidate_domains.items():
        if dom in seen_domains or _is_pre_fetch_blocked(dom):
            continue
        if dom in cache and str(cache[dom].get("status", "")).startswith("Aktif") and (cache[dom].get("primary_email") or cache[dom].get("phone")):
            hit = dict(cache[dom])
            _merge_firm_into_cache(cache, hit, sector_tag, inferred)
            hit.update({"item": item_name, "source": f"{sl} (Önbellek)", "match_type": mt})
            seen_domains.add(dom)
            if skip_domains is not None:
                skip_domains.add(dom)
            results.append(hit)
        else:
            to_fetch.append((dom, sl, mt))

    to_fetch = to_fetch[:max_sites]
    if to_fetch:
        total = len(to_fetch)
        with concurrent.futures.ThreadPoolExecutor(max_workers=20) as pool:
            futs = {}
            for dom, sl, mt in to_fetch:
                if skip_domains is not None:
                    skip_domains.add(dom)
                seen_domains.add(dom)
                fb = snippet_fallbacks.get(dom, {})
                futs[
                    pool.submit(
                        inspect_company_website,
                        dom,
                        fb.get("company_hint", ""),
                        False,
                        DEFAULT_TIMEOUT,
                        fb.get("emails"),
                        fb.get("phones"),
                        relevance_profile,
                    )
                ] = (dom, sl, mt)
            done = 0
            for fut in concurrent.futures.as_completed(futs):
                done += 1
                dom, sl, mt = futs[fut]
                if progress_cb:
                    progress_cb(f"{item_name} ({round_num}. Tur): {done}/{total} yeni site tarandı ({dom})")
                info = fut.result()
                fb = snippet_fallbacks.get(dom)
                if not info.get("status", "").startswith("Aktif"):
                    # Spam, yabancı site, alakasız sektör veya ücretsiz mail elemesi ise kesin reddet
                    if info.get("rejected"):
                        continue
                    # Yalnızca erişilemeyen ama arama özetinde kurumsal/TR sinyalli olanları kurtar
                    if fb and (fb["emails"] or fb["phones"]):
                        snip_surface = f"{fb.get('title', '')} {fb.get('company_hint', '')}"
                        if SPAM_BETTING_RE.search(snip_surface) or not _is_page_relevant(snip_surface, relevance_profile):
                            continue
                        prim, ordered = rank_emails(fb["emails"], dom)
                        has_corp_em = any(em.split("@")[-1] not in FREE_EMAIL_PROVIDERS for em in ordered)
                        if (has_corp_em or fb["phones"]) and (prim or fb["phones"]):
                            info = {
                                "company_name": _extract_clean_company_name(
                                    og_site_name="",
                                    copyright_cands=[],
                                    title=fb.get("title", ""),
                                    final_dom=dom,
                                    primary_email=prim,
                                    company_hint=fb.get("company_hint", ""),
                                ),
                                "domain": dom,
                                "website": fb["url"] or f"https://{dom}",
                                "primary_email": prim,
                                "all_emails": ordered,
                                "phone": " / ".join(_dedupe_phones(fb["phones"])[:2]),
                                "status": "Aktif (Arama Özetinden)",
                                "last_year": "-",
                                "notes": "Arama özetinden kurtarıldı",
                                "product_groups": ", ".join(inferred),
                            }
                        else:
                            continue
                    else:
                        continue
                if not info.get("primary_email") and not info.get("phone"):
                    continue
                info.update({
                    "item": item_name, "source": sl, "match_type": mt,
                    "sector": sector_tag, "last_checked": dt.date.today().isoformat(),
                })
                _merge_firm_into_cache(cache, info, sector_tag, inferred)
                results.append(info)

    tender_contractors = _resolve_tender_contractors(
        tender_hits, item_name, sector_tag, inferred, cache, progress_cb=progress_cb, timeout=DEFAULT_TIMEOUT
    )
    for tc in tender_contractors:
        tdom = str(tc.get("domain", "")).lower()
        if tdom and tdom not in seen_domains:
            seen_domains.add(tdom)
            if skip_domains is not None:
                skip_domains.add(tdom)
            _merge_firm_into_cache(cache, tc, sector_tag, inferred)
            results.append(tc)

    return results


def collect_bcc_emails(rows: list[dict[str, Any]], only_selected: bool = False) -> list[str]:
    emails: list[str] = []
    seen: set[str] = set()
    for r in rows:
        if only_selected and not r.get("selected", True):
            continue
        prim = (r.get("primary_email") or "").strip().lower()
        cands = ([prim] if prim else []) + [str(e).strip().lower() for e in r.get("all_emails", []) if str(e).strip()]
        for em in cands:
            if not em or em == "-" or em.endswith(".kep.tr") or em in seen:
                continue
            seen.add(em)
            emails.append(em)
    return emails


def build_excel_bytes(rows: list[dict[str, Any]]) -> bytes:
    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "Tedarikçiler"
    out_rows = [
        [
            r.get("item_name") or r.get("item", ""),
            r.get("company_name", ""),
            r.get("sector") or r.get("product_groups", ""),
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
    _write_sheet(ws_main, OUTPUT_COLUMNS, out_rows)
    ws_mails = wb.create_sheet(title="Mail Listesi (Alt Alta)")
    mail_rows = []
    seen_m: set[str] = set()
    for r in rows:
        prim = (r.get("primary_email") or "").strip().lower()
        for em in ([prim] if prim else []) + [str(e).strip().lower() for e in r.get("all_emails", [])]:
            if em and not em.endswith(".kep.tr") and em not in seen_m:
                seen_m.add(em)
                mail_rows.append([em, "Öncelikli" if em == prim else "Alternatif", r.get("item_name") or r.get("item", ""), r.get("company_name", "")])
    _write_sheet(ws_mails, ["E-posta (Kime/BCC'ye Yapıştır)", "Öncelik", "Kalem", "Firma Adı"], mail_rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def dataframe_to_utf8_bom_csv(df: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    df.to_csv(buf, index=False, sep=";", quoting=csv.QUOTE_MINIMAL)
    return buf.getvalue().encode("utf-8-sig")

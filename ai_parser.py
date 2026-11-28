"""Gemini API destekli İstem Formu ve Dijital Teklif Mektubu okuyucu.
PDF, Word (.docx), Excel (.xlsx/.xlsm/.csv) ve Görsel (.png/.jpg/.jpeg/.webp) dosyalarını ayrıştırır.
"""

from __future__ import annotations

import csv
import io
import json
import re
from pathlib import Path
from typing import Any

import openpyxl

BOILERPLATE_LINE_RE = re.compile(
    r"^(t\.?\s*c\.?|sağlık bakanlığı|devlet hastanesi|başhekim|satın alma|talep formu|"
    r"tarih|sayı|konu|imza|onay|adı soyadı|unvan|birim sorumlusu|not\s*:|açıklama\s*:)",
    re.IGNORECASE,
)
_SERVICE_KEYWORDS = (
    "bakım", "bakim", "onarım", "onarim", "kalibrasyon", "servis", "kiralama",
    "hizmet", "danışmanlık", "egitim", "eğitim", "temizlik", "güvenlik",
    "nakliye", "taşıma", "montaj", "periyodik", "ölçüm", "analiz", "yazılım",
)


def tr_lower(text: str) -> str:
    return str(text or "").replace("İ", "i").replace("I", "ı").lower()


def detect_procurement_type(items: list[str]) -> str:
    joined = tr_lower(" ".join(items))
    return "Hizmet" if any(kw in joined for kw in _SERVICE_KEYWORDS) else "Malzeme"

GEMINI_DOC_MODELS = (
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.1-flash-lite",
    "gemini-2.5-flash",
)

IMAGE_MIME_MAP = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}


def parse_tr_number(val: Any, default: float | None = None) -> float | None:
    """Türkçe veya uluslararası sayı formatını (1.250,50 / 1250.50 / 1.250 TL) float'a çevirir."""
    if val is None or isinstance(val, bool):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return default
    s = re.sub(r"(?i)\b(?:tl|try|usd|eur|₺|\$|€|kdv\s*hariç)\b", "", s)
    s = re.sub(r"[^\d.,\-]", "", s)
    if not s:
        return default
    # Hem nokta hem virgül varsa son görülen ondalık ayırıcıdır
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 4:
            s = parts[0].replace(".", "") + "." + parts[1]
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        # 1.250 veya 12.500 gibi 3 basamaklı binlik ayırıcıysa noktayı kaldır
        if len(parts) >= 2 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
    try:
        return float(s)
    except ValueError:
        return default


def _extract_office_text(filename: str, raw_bytes: bytes) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".docx":
        import docx

        doc = docx.Document(io.BytesIO(raw_bytes))
        blocks: list[str] = []
        for p in doc.paragraphs:
            t = p.text.strip()
            if t:
                blocks.append(t)
        for t_idx, tbl in enumerate(doc.tables, start=1):
            blocks.append(f"\n[TABLO {t_idx}]")
            for row in tbl.rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(cells):
                    blocks.append(" | ".join(cells))
        return "\n".join(blocks)

    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
        lines: list[str] = []
        for sheet_name in wb.sheetnames[:3]:
            ws = wb[sheet_name]
            lines.append(f"[SAYFA: {sheet_name}]")
            for row in ws.iter_rows(values_only=True):
                vals = [str(c).strip() if c is not None else "" for c in row]
                if any(vals):
                    lines.append(" | ".join(vals))
        wb.close()
        return "\n".join(lines)

    if ext == ".csv":
        text = raw_bytes.decode("utf-8-sig", errors="replace")
        lines = text.splitlines()
        delim = ";" if "\n".join(lines[:20]).count(";") > "\n".join(lines[:20]).count(",") else ","
        return "\n".join(" | ".join(c.strip() for c in r) for r in csv.reader(lines, delimiter=delim) if any(c.strip() for c in r))

    if ext == ".pdf":
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(raw_bytes))
        return "\n".join((p.extract_text() or "").strip() for p in reader.pages if (p.extract_text() or "").strip())

    if ext == ".txt":
        return raw_bytes.decode("utf-8-sig", errors="replace")

    return ""


def _build_gemini_parts(filename: str, raw_bytes: bytes, prompt_text: str) -> list[Any]:
    from google.genai import types

    ext = Path(filename).suffix.lower()
    if ext in IMAGE_MIME_MAP:
        return [
            types.Part.from_bytes(data=raw_bytes, mime_type=IMAGE_MIME_MAP[ext]),
            prompt_text,
        ]
    if ext == ".pdf":
        extracted = _extract_office_text(filename, raw_bytes)
        parts: list[Any] = [types.Part.from_bytes(data=raw_bytes, mime_type="application/pdf")]
        if extracted.strip():
            parts.append(f"PDF metin katmanı özeti:\n{extracted[:18000]}\n\n{prompt_text}")
        else:
            parts.append(prompt_text)
        return parts

    extracted = _extract_office_text(filename, raw_bytes)
    return [f"BELGE İÇERİĞİ ({filename}):\n{extracted[:25000]}\n\n{prompt_text}"]


def _call_gemini_json_strict(api_key: str, parts: list[Any], timeout_ms: int = 30000) -> tuple[dict[str, Any] | None, str]:
    """Fail-Closed JSON çağrısı: başarılıysa (data, ''), hata durumunda (None, gerçek_hata_mesajı) döner."""
    if not api_key:
        return None, "GEMINI_API_KEY tanımlı değil."
    try:
        from google import genai
        from google.genai import types
    except ImportError as exc:
        return None, f"google-genai paketi yüklenemedi: {exc}"

    client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=timeout_ms))
    last_err = "Bilinmeyen Gemini API hatası."
    for model_name in GEMINI_DOC_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=parts,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1),
            )
            txt = (resp.text or "").strip()
            if not txt:
                last_err = f"{model_name}: Boş yanıt döndü."
                continue
            if txt.startswith("```"):
                txt = re.sub(r"^```(?:json)?\s*|\s*```$", "", txt, flags=re.DOTALL).strip()
            data = json.loads(txt)
            if isinstance(data, dict):
                data["_model"] = model_name
                return data, ""
            last_err = f"{model_name}: JSON nesnesi yerine beklenmeyen tip döndü."
        except Exception as exc:
            last_err = f"{model_name}: {exc}"
            continue
    return None, last_err


def _call_gemini_json(api_key: str, parts: list[Any]) -> dict[str, Any] | None:
    data, _ = _call_gemini_json_strict(api_key, parts)
    return data


# ---------------------------------------------------------------------------
# Adım 1: İstem Formu Ayrıştırma
# ---------------------------------------------------------------------------

def _fallback_parse_demand_local(filename: str, raw_bytes: bytes) -> dict[str, Any]:
    """Gemini erişilemediğinde Word/Excel/CSV/PDF metninden temel kalem çıkarımı yapar."""
    ext = Path(filename).suffix.lower()
    items: list[dict[str, Any]] = []
    if ext in IMAGE_MIME_MAP:
        return {
            "alim_adi": Path(filename).stem,
            "alim_turu": "Malzeme",
            "kalemler": [],
            "source": "Yerel",
            "warning": "Görsel dosyaları okumak için geçerli bir GEMINI_API_KEY gereklidir.",
        }

    text = _extract_office_text(filename, raw_bytes)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    idx_counter = 1
    for line in lines:
        if line.startswith("[") and line.endswith("]"):
            continue
        if "|" in line:
            cols = [c.strip() for c in line.split("|")]
            if len(cols) >= 2:
                joined_low = tr_lower(" ".join(cols))
                if any(h in joined_low for h in ("malzeme adı", "hizmet adı", "sıra no", "birim", "miktar")) and len(cols) <= 6:
                    continue
                # İlk sütun sıra no ise ikinci sütun isim olabilir
                name = cols[1] if (re.fullmatch(r"\d+[.)]?", cols[0]) and len(cols) > 1) else cols[0]
                if not name or len(name) < 2 or BOILERPLATE_LINE_RE.match(name):
                    continue
                unit = "Adet"
                qty = 1.0
                for c in cols[2:]:
                    num = parse_tr_number(c)
                    if num is not None and num > 0:
                        qty = num
                    elif 1 <= len(c) <= 15 and not c.isdigit():
                        unit = c
                items.append({"sira_no": idx_counter, "name": name, "unit": unit, "quantity": qty})
                idx_counter += 1
        elif re.match(r"^\d+[.)\-]\s+", line):
            clean = re.sub(r"^\d+[.)\-]\s+", "", line).strip()
            if len(clean) >= 3 and not BOILERPLATE_LINE_RE.match(clean):
                items.append({"sira_no": idx_counter, "name": clean, "unit": "Adet", "quantity": 1.0})
                idx_counter += 1

    alim_turu = detect_procurement_type([it["name"] for it in items]) if items else "Malzeme"
    return {
        "alim_adi": Path(filename).stem.replace("_", " ").strip() or "Yeni Alım Dosyası",
        "alim_turu": alim_turu,
        "kalemler": items[:100],
        "source": "Yerel Ayrıştırıcı",
    }


def parse_demand_form(filename: str, raw_bytes: bytes, api_key: str) -> dict[str, Any]:
    """Satın Alma İstem Formundan alım adı, alım türü (Malzeme/Hizmet) ve kalem listesini çıkarır."""
    prompt = (
        "Sen bir kamu hastanesi satın alma uzmanısın. Ekteki Satın Alma İstem Formunu / Talep Belgesini dikkatlice incele.\n"
        "Belgeden alımın genel adını, alım türünü ve talep edilen tüm kalemleri çıkar.\n"
        "Yanıtı YALNIZCA aşağıdaki şemaya uygun geçerli JSON olarak ver:\n"
        "{\n"
        '  "alim_adi": "Alımı özetleyen kısa ve net başlık (örn. Biyomedikal Cihaz Bakım Onarımı veya Kardiyoloji Sarf Malzeme Alımı)",\n'
        '  "alim_turu": "Malzeme" veya "Hizmet",\n'
        '  "kalemler": [\n'
        "    {\n"
        '      "sira_no": 1,\n'
        '      "name": "Malzeme veya hizmetin adı ve varsa kısa açıklaması",\n'
        '      "unit": "Adet / Kutu / Paket / Kg / Litre / Ay / Hizmet / Set vb.",\n'
        '      "quantity": 10\n'
        "    }\n"
        "  ]\n"
        "}\n"
        "Kurallar:\n"
        "- Tablo başlıklarını, hastane unvanını, imza bloklarını ve boş satırları kalem olarak ekleme.\n"
        "- alim_turu değeri yalnızca 'Malzeme' veya 'Hizmet' olabilir.\n"
        "- quantity mutlaka sayısal değer olsun (belirtilmemişse 1 yaz).\n"
    )
    if api_key:
        try:
            parts = _build_gemini_parts(filename, raw_bytes, prompt)
            data = _call_gemini_json(api_key, parts)
            if data and isinstance(data.get("kalemler"), list) and data["kalemler"]:
                cleaned_items: list[dict[str, Any]] = []
                for idx, raw_it in enumerate(data["kalemler"], start=1):
                    if not isinstance(raw_it, dict):
                        continue
                    name = str(raw_it.get("name") or raw_it.get("ad") or "").strip()
                    if not name:
                        continue
                    qty = parse_tr_number(raw_it.get("quantity"), 1.0)
                    if qty is None or qty <= 0:
                        qty = 1.0
                    sira = raw_it.get("sira_no")
                    try:
                        sira_int = int(sira) if sira is not None else idx
                    except (TypeError, ValueError):
                        sira_int = idx
                    cleaned_items.append({
                        "sira_no": sira_int,
                        "name": name,
                        "unit": str(raw_it.get("unit") or "Adet").strip() or "Adet",
                        "quantity": qty,
                    })
                if cleaned_items:
                    raw_type = str(data.get("alim_turu") or "").strip()
                    alim_turu = "Hizmet" if "hizmet" in tr_lower(raw_type) else (
                        "Malzeme" if "malzeme" in tr_lower(raw_type) else detect_procurement_type([x["name"] for x in cleaned_items])
                    )
                    return {
                        "alim_adi": str(data.get("alim_adi") or Path(filename).stem).strip(),
                        "alim_turu": alim_turu,
                        "kalemler": cleaned_items,
                        "source": f"Gemini ({data.get('_model', 'flash')})",
                    }
        except Exception:
            pass

    return _fallback_parse_demand_local(filename, raw_bytes)


# ---------------------------------------------------------------------------
# Adım 4: Dijital Teklif Mektubu Okuma
# ---------------------------------------------------------------------------

def parse_offer_document(
    filename: str,
    raw_bytes: bytes,
    case_items: list[dict[str, Any]],
    api_key: str,
) -> dict[str, Any]:
    """Teklif mektubundan (PDF, Word, Excel, Görsel) firma adını ve kalem bazlı birim fiyatları çıkarır."""
    items_desc = "\n".join(
        f"- item_id={it['id']} | sıra_no={it['sira_no']} | ad={it['name']} | birim={it['unit']} | miktar={it['quantity']}"
        for it in case_items
    )
    prompt = (
        "Sen bir kamu hastanesi satın alma komisyonu uzmanısın. Ekteki tedarikçi Fiyat Teklif Mektubunu incele.\n"
        "Aşağıdaki ihtiyaç listesi kalemleriyle teklif belgesindeki satırları eşleştir ve her kalem için KDV HARİÇ BİRİM FİYATINI çıkar.\n"
        f"İHTİYAÇ LİSTESİ:\n{items_desc}\n\n"
        "Yanıtı YALNIZCA aşağıdaki şemaya uygun geçerli JSON olarak ver:\n"
        "{\n"
        '  "firma_adi": "Teklifi veren firmanın tam adı (kaşe/antetten veya imzadan)",\n'
        '  "fiyatlar": [\n'
        '    {"item_id": 1, "sira_no": 1, "birim_fiyat": 1250.50}\n'
        "  ]\n"
        "}\n"
        "Kurallar:\n"
        "- Eğer belgede sadece toplam tutar verilmişse ve birim fiyat yoksa toplam tutarı miktara bölerek birim fiyatı hesapla; mümkünse doğrudan Birim Fiyat sütununu al.\n"
        "- Firma belirli bir kaleme teklif vermediyse (boş, 'teklif yok', '---' ise) birim_fiyat olarak null yaz.\n"
        "- Sayı formatını nokta ondalıklı saf sayı olarak ver (ör. 1450.75).\n"
    )

    if api_key:
        try:
            parts = _build_gemini_parts(filename, raw_bytes, prompt)
            data = _call_gemini_json(api_key, parts)
            if data and (data.get("firma_adi") or data.get("fiyatlar")):
                item_by_id = {int(it["id"]): it for it in case_items}
                item_by_sira = {int(it["sira_no"]): it for it in case_items}
                prices_by_item_id: dict[int, float | None] = {}
                for row in (data.get("fiyatlar") or []):
                    if not isinstance(row, dict):
                        continue
                    iid = row.get("item_id")
                    target_id: int | None = None
                    try:
                        if iid is not None and int(iid) in item_by_id:
                            target_id = int(iid)
                    except (TypeError, ValueError):
                        pass
                    if target_id is None:
                        try:
                            sr = int(row.get("sira_no"))
                            if sr in item_by_sira:
                                target_id = int(item_by_sira[sr]["id"])
                        except (TypeError, ValueError):
                            pass
                    if target_id is not None:
                        val = parse_tr_number(row.get("birim_fiyat"), None)
                        prices_by_item_id[target_id] = val if (val is not None and val > 0) else None
                return {
                    "firma_adi": str(data.get("firma_adi") or Path(filename).stem).strip(),
                    "prices": prices_by_item_id,
                    "source": f"Gemini ({data.get('_model', 'flash')})",
                }
        except Exception:
            pass

    # Yerel Word/Excel yedek ayrıştırma
    ext = Path(filename).suffix.lower()
    prices_by_item_id = {}
    if ext in (".docx", ".xlsx", ".xlsm", ".csv"):
        text = _extract_office_text(filename, raw_bytes)
        lines = [l for l in text.splitlines() if "|" in l]
        for it in case_items:
            sira_str = str(it["sira_no"])
            it_name_low = tr_lower(it["name"])[:18]
            for ln in lines:
                cols = [c.strip() for c in ln.split("|")]
                if len(cols) >= 5 and (cols[0] == sira_str or it_name_low in tr_lower(ln)):
                    cand = parse_tr_number(cols[-2], None)
                    if cand is not None and cand > 0:
                        prices_by_item_id[int(it["id"])] = cand
                        break
    return {
        "firma_adi": Path(filename).stem.replace("_", " ").strip(),
        "prices": prices_by_item_id,
        "source": "Yerel Ayrıştırıcı",
    }


# ---------------------------------------------------------------------------
# Genel B2B Sektör ve Sorgu Üretici (Tüm Malzeme/Hizmet Sektörleri)
# ---------------------------------------------------------------------------

def build_b2b_sector_queries(item_name: str, sector_name: str, synonyms: list[str] | None = None) -> list[str]:
    """Perakende yerine doğrudan toptancı, bölgesel (Balıkesir/İzmir/Bursa) ve kamu yüklenicisi B2B sorguları üretir."""
    clean_item = re.sub(r"\([^)]*\)", " ", str(item_name or "")).strip()
    clean_item = re.sub(r"\s+", " ", re.sub(r"\b(için|icin|grubu|tipi|adet|kutu)\b", " ", clean_item, flags=re.IGNORECASE)).strip() or item_name.strip()
    kw = (synonyms[0].strip() if synonyms and synonyms[0].strip() else clean_item)
    sec = (sector_name or "Kurumsal Tedarik").split(",")[0].strip()
    return list(dict.fromkeys([
        f'"{sec}" toptan satış "Ltd. Şti." OR "A.Ş." iletişim',
        f'"{kw}" "yüklenici" OR "ihale sonucu" sözleşme',
        f'"{sec}" kurumsal tedarik teklif e-posta',
        f'"{sec}" Balıkesir OR İzmir OR Bursa toptan iletişim',
        f'"{clean_item}" toptan satış tedarikçi iletişim "@"',
    ]))


def generate_sector_and_b2b_queries(
    item_name: str,
    api_key: str,
    fallback_sector: str = "Kurumsal Tedarik",
) -> dict[str, Any]:
    """Gemini ile kalemin gerçek ticari sektörünü belirler ve toptancı/kamu tedarikçisi B2B sorguları üretir."""
    fallback_queries = build_b2b_sector_queries(item_name, fallback_sector)
    fallback_result = {
        "aciklama": (
            f"{item_name} kalemi kurum birimlerinin doğrudan temin ve piyasa araştırması kapsamında ihtiyaç duyduğu malzeme/hizmettir. "
            "Türkiye genelinde ve çevre illerde faaliyet gösteren toptan tedarikçi ve yetkili firmalardan fiyat teklifi toplanır."
        ),
        "es_anlamlilar": [item_name],
        "ust_kategori": fallback_sector,
        "markalar": [],
        "b2b_sorgular": fallback_queries,
        "source": "Yerel Analiz",
    }
    if not api_key:
        return fallback_result

    prompt = (
        f"Bir kamu hastanesi ve idari satın alma birimi için şu kalemi analiz et: '{item_name}'.\n"
        "Sistem sadece medikal değil hırdavat, sıhhi tesisat, elektrik, kırtasiye, bilişim, mobilya, kimya, ilaç, teknik servis gibi TÜM sektörlerde çalışır.\n"
        "Yanıtı YALNIZCA aşağıdaki anahtarlara sahip geçerli bir JSON nesnesi olarak ver:\n"
        "- ust_kategori: Bu kalemin ait olduğu GERÇEK ticari sektör adı (ör. 'Lavabo Bataryası' -> 'Sıhhi Tesisat ve Yapı Malzemeleri', 'Toner' -> 'Büro Makineleri ve Sarf Malzemeleri', 'EKG Cihazı' -> 'Tıbbi Cihaz ve Medikal Sistemler').\n"
        "- aciklama: Tam olarak 2 cümle; ürünün/hizmetin ne olduğunu ve hangi ticari sektörde tedarik edildiğini açıkla.\n"
        "- es_anlamlilar: Türkiye piyasasında toptancıların kullandığı 3-5 adet eş anlamlı / teknik arama terimi.\n"
        "- markalar: Türkiye'de bu alanda bilinen 2-4 üretici veya ana marka adı listesi (yoksa boş liste).\n"
        "- b2b_sorgular: Perakende müşterileri değil DOĞRUDAN toptancıları, bölgesel (Balıkesir/İzmir/Bursa) tedarikçileri ve kamu ihale yüklenicilerini hedefleyen TAM OLARAK 5 B2B arama sorgusu üret:\n"
        '  1) "{ust_kategori}" toptan satış "Ltd. Şti." OR "A.Ş." iletişim\n'
        '  2) "{kalem_anahtar_kelime}" "yüklenici" OR "ihale sonucu" sözleşme\n'
        '  3) "{ust_kategori}" kurumsal tedarik teklif e-posta\n'
        '  4) "{ust_kategori}" Balıkesir OR İzmir OR Bursa toptan iletişim\n'
        '  5) "{kalem_anahtar_kelime}" toptan satış tedarikçi iletişim "@"\n'
    )
    data = _call_gemini_json(api_key, [prompt])
    if data and data.get("ust_kategori"):
        sec = str(data.get("ust_kategori") or fallback_sector).strip()
        syns = [str(x).strip() for x in (data.get("es_anlamlilar") or []) if str(x).strip()][:5]
        raw_b2b = [str(x).strip() for x in (data.get("b2b_sorgular") or []) if str(x).strip()]
        structured = build_b2b_sector_queries(item_name, sec, syns)
        merged_b2b = list(dict.fromkeys(raw_b2b[:3] + structured))[:6]
        return {
            "aciklama": str(data.get("aciklama") or fallback_result["aciklama"]).strip(),
            "es_anlamlilar": syns or [item_name],
            "ust_kategori": sec,
            "markalar": [str(x).strip() for x in (data.get("markalar") or []) if str(x).strip()][:4],
            "b2b_sorgular": merged_b2b,
            "source": f"{data.get('_model', 'gemini-flash')}",
            "schema_v2": True,
        }
    return fallback_result


# ---------------------------------------------------------------------------
# MKYS Geçmiş Alım Excel Ayrıştırıcı
# ---------------------------------------------------------------------------

_MKYS_ITEM_HINTS = (
    "malzeme/hizmet", "malzeme tanımı", "malzeme tanimi", "malzeme adı", "malzeme adi",
    "hizmet adı", "hizmet adi", "taşınır adı", "tasinir adi", "kalem adı", "kalem adi",
    "malzeme cinsi", "ürün adı", "malzeme", "hizmet", "kalem", "cinsi",
)
_MKYS_FIRM_HINTS = (
    "tedarikçi firma", "tedarikci firma", "firma ünvanı", "firma unvanı", "firma unvani",
    "tedarikçi ünvanı", "tedarikci unvani", "firma adı", "firma adi", "satıcı firma",
    "tedarikçi", "tedarikci", "yüklenici", "yuklenici", "satıcı", "firma", "ünvan", "unvan",
)
_MKYS_PUBLIC_RE = re.compile(
    r"\b(DEVLET MALZEME OFİSİ|DMO|HASTANESİ|İL SAĞLIK|BAKANLIĞI|BELEDİYESİ|ÜNİVERSİTESİ|SAYMANLIĞI)\b",
    re.IGNORECASE,
)


def _read_tabular_rows(filename: str, raw_bytes: bytes) -> list[list[Any]]:
    ext = Path(filename).suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, read_only=True)
        rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        return rows
    if ext == ".csv":
        text = raw_bytes.decode("utf-8-sig", errors="replace")
        lines = text.splitlines()
        delim = ";" if "\n".join(lines[:20]).count(";") > "\n".join(lines[:20]).count(",") else ","
        return list(csv.reader(lines, delimiter=delim))
    return []


def parse_mkys_excel(
    filename: str,
    raw_bytes: bytes,
    exclude_public: bool = True,
    item_col_override: int | None = None,
    firm_col_override: int | None = None,
) -> dict[str, Any]:
    """MKYS İş Zekası Excel/CSV dosyasından Malzeme/Hizmet Adı ve Tedarikçi Firma Ünvanı kayıtlarını çıkarır."""
    rows = _read_tabular_rows(filename, raw_bytes)
    non_empty = [r for r in rows if any(str(c or "").strip() for c in r)]
    if not non_empty:
        return {"error": "Dosyada okunabilir tablo satırı bulunamadı.", "records": [], "headers": []}

    best_idx = 0
    best_score = -1
    best_item_col: int | None = None
    best_firm_col: int | None = None

    for i, r in enumerate(non_empty[:40]):
        if sum(1 for c in r if str(c or "").strip()) < 2:
            continue
        lows = [tr_lower(str(c or "").strip()) for c in r]
        it_c = None
        fm_c = None
        for h in _MKYS_ITEM_HINTS:
            for idx, col in enumerate(lows):
                if h in col and not any(ex in col for ex in ("birim adı", "birim adi", "hastane", "kurum", "bölüm")):
                    it_c = idx
                    break
            if it_c is not None:
                break
        for h in _MKYS_FIRM_HINTS:
            for idx, col in enumerate(lows):
                if idx == it_c:
                    continue
                if h in col:
                    fm_c = idx
                    break
            if fm_c is not None:
                break
        sc = (3 if it_c is not None else 0) + (3 if fm_c is not None else 0)
        if sc > best_score:
            best_score = sc
            best_idx = i
            best_item_col = it_c
            best_firm_col = fm_c

    headers = [str(c or f"Sütun {idx + 1}").strip() or f"Sütun {idx + 1}" for idx, c in enumerate(non_empty[best_idx])]
    it_col = item_col_override if item_col_override is not None else best_item_col
    fm_col = firm_col_override if firm_col_override is not None else best_firm_col

    if it_col is None or fm_col is None:
        return {
            "headers": headers,
            "detected_item_col": it_col,
            "detected_firm_col": fm_col,
            "records": [],
            "skipped_public": 0,
            "total_rows": max(0, len(non_empty) - best_idx - 1),
            "needs_mapping": True,
        }

    # Opsiyonel sütunlar: Birim, Miktar, Birim Fiyat, Tarih, Sektör
    low_headers = [tr_lower(h) for h in headers]
    unit_col = next((i for i, h in enumerate(low_headers) if "birim" in h and "fiyat" not in h and "adı" not in h), None)
    qty_col = next((i for i, h in enumerate(low_headers) if "miktar" in h), None)
    price_col = next((i for i, h in enumerate(low_headers) if "birim fiyat" in h or "fiyat" in h), None)
    date_col = next((i for i, h in enumerate(low_headers) if "tarih" in h), None)
    sec_col = next((i for i, h in enumerate(low_headers) if "sektör" in h or "sektor" in h or "grup" in h), None)

    records: list[dict[str, Any]] = []
    skipped_public = 0
    seen_pairs: set[tuple[str, str]] = set()

    for r in non_empty[best_idx + 1:]:
        it_val = str(r[it_col] or "").strip() if it_col < len(r) else ""
        fm_val = str(r[fm_col] or "").strip() if fm_col < len(r) else ""
        if not it_val or not fm_val or len(it_val) < 2 or len(fm_val) < 2:
            continue
        if BOILERPLATE_LINE_RE.match(it_val) or BOILERPLATE_LINE_RE.match(fm_val):
            continue
        if exclude_public and _MKYS_PUBLIC_RE.search(fm_val.upper()):
            skipped_public += 1
            continue
        pair_key = (tr_lower(it_val), tr_lower(fm_val))
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)

        records.append({
            "item_name": it_val,
            "supplier_name": fm_val,
            "sector": str(r[sec_col] or "").strip() if (sec_col is not None and sec_col < len(r)) else "",
            "unit": str(r[unit_col] or "").strip() if (unit_col is not None and unit_col < len(r)) else "",
            "quantity": parse_tr_number(r[qty_col], None) if (qty_col is not None and qty_col < len(r)) else None,
            "unit_price": parse_tr_number(r[price_col], None) if (price_col is not None and price_col < len(r)) else None,
            "purchase_date": str(r[date_col] or "").strip() if (date_col is not None and date_col < len(r)) else "",
        })

    return {
        "headers": headers,
        "detected_item_col": it_col,
        "detected_firm_col": fm_col,
        "records": records,
        "skipped_public": skipped_public,
        "total_rows": max(0, len(non_empty) - best_idx - 1),
        "needs_mapping": False,
    }


# ---------------------------------------------------------------------------
# Fail-Closed Kayıt Öncesi Yapay Zeka Tampon Bölgesi & Otomatik Havuz Denetimi
# ---------------------------------------------------------------------------

def verify_suppliers_with_ai(
    candidates: list[dict[str, Any]],
    item_name: str,
    api_key: str,
    batch_size: int = 20,
) -> dict[str, Any]:
    """20'şerli paketler halinde aday tedarikçileri web içerikleriyle Gemini'de denetler.
    FAIL-CLOSED: Herhangi bir API/kota/zaman aşımı/JSON hatasında asla onaysız aday döndürmez (verified=[] + error).
    """
    if not candidates:
        return {"verified": [], "rejected_count": 0, "error": ""}
    if not api_key:
        return {
            "verified": [],
            "rejected_count": 0,
            "error": "GEMINI_API_KEY tanımlı değil. Fail-Closed güvenlik kuralı gereği onaysız adaylar kaydedilmedi.",
        }

    verified_all: list[dict[str, Any]] = []
    rejected_total = 0

    for start in range(0, len(candidates), batch_size):
        batch = candidates[start : start + batch_size]
        payload = [
            {
                "idx": idx,
                "domain": str(c.get("domain") or "").strip().lower(),
                "company_name": str(c.get("company_name") or "").strip(),
                "website": str(c.get("website") or "").strip(),
                "title": str(c.get("site_title") or "").strip()[:160],
                "meta": str(c.get("site_meta") or "").strip()[:220],
                "content_excerpt": str(c.get("site_excerpt") or c.get("notes") or "").strip()[:700],
            }
            for idx, c in enumerate(batch)
        ]

        prompt = (
            f"Kamu kurumu/hastanesi satın alma birimi için '{item_name}' kalemini tedarik edebilecek gerçek firmaları denetle.\n"
            "Aşağıdaki aday listesinin her biri için kendi web sitesi içeriğini (title, meta, content_excerpt) inceleyerek karar ver.\n"
            f"ADAYLAR:\n{json.dumps(payload, ensure_ascii=False)}\n\n"
            "Yanıtı YALNIZCA aşağıdaki şemada geçerli JSON olarak ver:\n"
            "{\n"
            '  "results": [\n'
            "    {\n"
            '      "idx": 0,\n'
            '      "domain": "ornek.com.tr",\n'
            '      "is_valid_supplier": true,\n'
            '      "real_sector": "Firmanın kendi içeriğine göre gerçek ticari faaliyet alanı (ör. Yapı Market, Ecza Deposu, Tıbbi Sarf, Sıhhi Tesisat)",\n'
            '      "clean_name": "Ana Sayfa veya ürün listesi çöpleri temizlenmiş resmi şirket ünvanı"\n'
            "    }\n"
            "  ]\n"
            "}\n"
            "KESİN KURALLAR:\n"
            f"- is_valid_supplier: Bu firma aranan '{item_name}' kalemini resmi/faturalı olarak tedarik edebilecek gerçek bir satıcı, toptancı, üretici, distribütör veya ecza deposu ise true; DEĞİLSE false.\n"
            "- Haber siteleri, telekom operatörleri, bahis/kumar/kredi siteleri, prospektüs/sağlık-bilgi/blog/sözlük siteleri, ilan portalları ve farklı sektör firmaları için KESİNLİKLE is_valid_supplier: false ver.\n"
            "- clean_name içinde asla 'Ana Sayfa', 'İletişim', 'Ürünlerimiz', fiyat veya uzun ürün listesi olmasın.\n"
        )
        data, err = _call_gemini_json_strict(api_key, [prompt], timeout_ms=35000)
        if err or not data or not isinstance(data.get("results"), list):
            return {
                "verified": [],
                "rejected_count": 0,
                "error": f"Gemini AI doğrulama hatası (Fail-Closed): {err or 'Geçersiz JSON sonuç listesi'}",
            }

        decisions_by_idx: dict[int, dict[str, Any]] = {}
        decisions_by_dom: dict[str, dict[str, Any]] = {}
        for item in data["results"]:
            if not isinstance(item, dict):
                continue
            if "idx" in item:
                try:
                    decisions_by_idx[int(item["idx"])] = item
                except (TypeError, ValueError):
                    pass
            dom_k = str(item.get("domain") or "").strip().lower()
            if dom_k:
                decisions_by_dom[dom_k] = item

        for idx, cand in enumerate(batch):
            dom_k = str(cand.get("domain") or "").strip().lower()
            dec = decisions_by_idx.get(idx) or decisions_by_dom.get(dom_k)
            if not dec or dec.get("is_valid_supplier") is not True:
                rejected_total += 1
                continue

            clean_nm = str(dec.get("clean_name") or "").strip()
            real_sec = str(dec.get("real_sector") or "").strip()
            verified_cand = dict(cand)
            if clean_nm and len(clean_nm) >= 2:
                verified_cand["company_name"] = clean_nm[:80]
            if real_sec:
                verified_cand["sector"] = real_sec[:120]
            verified_cand["ai_verified"] = 1
            verified_all.append(verified_cand)

    return {
        "verified": verified_all,
        "rejected_count": rejected_total,
        "error": "",
    }


def sanitize_pool_batch_with_ai(
    pool_rows: list[dict[str, Any]],
    mkys_rows: list[dict[str, Any]],
    api_key: str,
    batch_size: int = 20,
) -> dict[str, Any]:
    """ai_verified = 0 olan mevcut havuz ve MKYS kayıtlarını sessizce Gemini ile denetler."""
    if not api_key:
        return {"pool_decisions": [], "mkys_decisions": [], "error": "GEMINI_API_KEY tanımlı değil."}

    pool_decisions: list[dict[str, Any]] = []
    mkys_decisions: list[dict[str, Any]] = []

    for start in range(0, len(pool_rows), batch_size):
        batch = pool_rows[start : start + batch_size]
        payload = [
            {
                "id": int(r["id"]),
                "company_name": str(r.get("company_name") or ""),
                "domain": str(r.get("domain") or ""),
                "website": str(r.get("website") or ""),
                "current_sector": str(r.get("sektor") or ""),
                "items_sold": str(r.get("items_sold_json") or "[]")[:300],
            }
            for r in batch
        ]
        prompt = (
            "Satın alma veritabanındaki mevcut kayıtları temizle. Her kayıt için gerçek bir ticari tedarikçi olup olmadığını belirle.\n"
            f"KAYITLAR:\n{json.dumps(payload, ensure_ascii=False)}\n\n"
            "Yanıtı YALNIZCA şu JSON şemasında ver:\n"
            '{"decisions": [{"id": 1, "keep": true, "real_sector": "Gerçek Faaliyet Alanı", "clean_name": "Temiz Şirket Ünvanı"}]}\n'
            "Kurallar:\n"
            "- Haber, medya, telekom, bahis, kredi, prospektüs/ilaç-bilgi siteleri, blog/sözlük, sosyal medya, pazaryeri veya spam kayıtlar için keep: false ver.\n"
            "- Gerçek satıcı, toptancı, üretici, distribütör, yapı market, ecza deposu veya servis firmaları için keep: true ver ve real_sector alanını sitenin/ünvanın gerçek faaliyet alanıyla doldur.\n"
        )
        data, err = _call_gemini_json_strict(api_key, [prompt], timeout_ms=30000)
        if err or not data or not isinstance(data.get("decisions"), list):
            return {"pool_decisions": pool_decisions, "mkys_decisions": mkys_decisions, "error": err or "Geçersiz JSON"}
        for d in data["decisions"]:
            if isinstance(d, dict) and "id" in d:
                pool_decisions.append({
                    "id": int(d["id"]),
                    "keep": d.get("keep") is True,
                    "real_sector": str(d.get("real_sector") or "").strip(),
                    "clean_name": str(d.get("clean_name") or "").strip(),
                })

    for start in range(0, len(mkys_rows), batch_size):
        batch = mkys_rows[start : start + batch_size]
        payload = [
            {
                "id": int(r["id"]),
                "supplier_name": str(r.get("supplier_name") or ""),
                "item_name": str(r.get("item_name") or ""),
                "current_sector": str(r.get("sektor") or ""),
            }
            for r in batch
        ]
        prompt = (
            "MKYS geçmiş alım kayıtlarındaki tedarikçi firmaları denetle ve her kaydın gerçek ticari sektörünü belirle.\n"
            f"KAYITLAR:\n{json.dumps(payload, ensure_ascii=False)}\n\n"
            "Yanıtı YALNIZCA şu JSON şemasında ver:\n"
            '{"decisions": [{"id": 1, "keep": true, "real_sector": "Gerçek Ticari Sektör"}]}\n'
            "Kurallar:\n"
            "- Kamu kurumu (DMO, hastane, bakanlık vb.), spam, anlamsız veya geçersiz firma adıysa keep: false ver.\n"
            "- Gerçek ticari firma ise keep: true ver ve satılan kaleme göre gerçek ticari sektörü yaz.\n"
        )
        data, err = _call_gemini_json_strict(api_key, [prompt], timeout_ms=30000)
        if err or not data or not isinstance(data.get("decisions"), list):
            return {"pool_decisions": pool_decisions, "mkys_decisions": mkys_decisions, "error": err or "Geçersiz JSON"}
        for d in data["decisions"]:
            if isinstance(d, dict) and "id" in d:
                mkys_decisions.append({
                    "id": int(d["id"]),
                    "keep": d.get("keep") is True,
                    "real_sector": str(d.get("real_sector") or "").strip(),
                })

    return {"pool_decisions": pool_decisions, "mkys_decisions": mkys_decisions, "error": ""}

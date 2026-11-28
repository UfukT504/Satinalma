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

from search_engine import BOILERPLATE_LINE_RE, detect_procurement_type, tr_lower

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


def _call_gemini_json(api_key: str, parts: list[Any]) -> dict[str, Any] | None:
    if not api_key:
        return None
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        return None

    client = genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=30000))
    for model_name in GEMINI_DOC_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=parts,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1),
            )
            txt = (resp.text or "").strip()
            if txt.startswith("```"):
                txt = re.sub(r"^```(?:json)?\s*|\s*```$", "", txt, flags=re.DOTALL).strip()
            data = json.loads(txt)
            if isinstance(data, dict):
                data["_model"] = model_name
                return data
        except Exception:
            continue
    return None


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

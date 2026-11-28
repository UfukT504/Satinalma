"""Yerel SQLite veritabanı katmanı — Dosya bazlı Satın Alma Yönetim Sistemi."""

from __future__ import annotations

import contextlib
import datetime as dt
import json
import sqlite3
from pathlib import Path
from typing import Any, Iterator

DB_PATH = Path(__file__).resolve().parent / "satin_alma.db"

DEFAULT_COMMISSION = {
    "members": [
        {"name": "", "title": "Satın Alma Komisyon Üyesi"},
        {"name": "", "title": "Satın Alma Komisyon Üyesi"},
        {"name": "", "title": "Satın Alma Komisyon Üyesi"},
    ],
    "authorizer_name": "",
    "authorizer_title": "Başhekim / Hastane Harcama Yetkilisi",
}


def _now_iso() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


@contextlib.contextmanager
def get_conn(db_path: Path | None = None) -> Iterator[sqlite3.Connection]:
    target = db_path or DB_PATH
    conn = sqlite3.connect(str(target), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA foreign_keys = ON;")
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, definition: str) -> None:
    cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def init_db(db_path: Path | None = None) -> None:
    with get_conn(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                procurement_type TEXT NOT NULL DEFAULT 'Malzeme',
                deadline_date TEXT NOT NULL DEFAULT '',
                deadline_time TEXT NOT NULL DEFAULT '',
                evaluation_mode TEXT NOT NULL DEFAULT 'Kalem Bazlı En Düşük',
                commission_json TEXT NOT NULL DEFAULT '[]',
                authorizer_name TEXT NOT NULL DEFAULT '',
                authorizer_title TEXT NOT NULL DEFAULT 'Başhekim / Hastane Harcama Yetkilisi',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS case_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL REFERENCES cases(id) ON DELETE CASCADE,
                sira_no INTEGER NOT NULL DEFAULT 1,
                name TEXT NOT NULL,
                unit TEXT NOT NULL DEFAULT 'Adet',
                quantity REAL NOT NULL DEFAULT 1,
                ai_intel_json TEXT NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS case_suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL REFERENCES cases(id) ON DELETE CASCADE,
                item_name TEXT NOT NULL DEFAULT '',
                company_name TEXT NOT NULL,
                domain TEXT NOT NULL DEFAULT '',
                website TEXT NOT NULL DEFAULT '',
                primary_email TEXT NOT NULL DEFAULT '',
                all_emails_json TEXT NOT NULL DEFAULT '[]',
                phone TEXT NOT NULL DEFAULT '',
                sector TEXT NOT NULL DEFAULT '',
                match_type TEXT NOT NULL DEFAULT '',
                source TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'Aktif',
                notes TEXT NOT NULL DEFAULT '',
                selected INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS case_offers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                case_id INTEGER NOT NULL REFERENCES cases(id) ON DELETE CASCADE,
                firm_name TEXT NOT NULL,
                source_type TEXT NOT NULL DEFAULT 'Manuel',
                source_file TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS case_offer_prices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                offer_id INTEGER NOT NULL REFERENCES case_offers(id) ON DELETE CASCADE,
                item_id INTEGER NOT NULL REFERENCES case_items(id) ON DELETE CASCADE,
                unit_price REAL
            );

            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS mkys_gecmis_alimlar (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_name TEXT NOT NULL,
                item_norm TEXT NOT NULL,
                supplier_name TEXT NOT NULL,
                supplier_norm TEXT NOT NULL,
                sektor TEXT NOT NULL DEFAULT '',
                unit TEXT NOT NULL DEFAULT '',
                quantity REAL,
                unit_price REAL,
                purchase_date TEXT NOT NULL DEFAULT '',
                ai_verified INTEGER NOT NULL DEFAULT 0,
                guven_puani INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );
            CREATE UNIQUE INDEX IF NOT EXISTS idx_mkys_alim_item_sup ON mkys_gecmis_alimlar(item_norm, supplier_norm);

            CREATE TABLE IF NOT EXISTS ana_firma_havuzu (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_name TEXT NOT NULL,
                company_norm TEXT NOT NULL,
                domain TEXT NOT NULL DEFAULT '',
                website TEXT NOT NULL DEFAULT '',
                primary_email TEXT NOT NULL DEFAULT '',
                all_emails_json TEXT NOT NULL DEFAULT '[]',
                phone TEXT NOT NULL DEFAULT '',
                sektor TEXT NOT NULL DEFAULT '',
                items_sold_json TEXT NOT NULL DEFAULT '[]',
                source TEXT NOT NULL DEFAULT '',
                last_checked TEXT NOT NULL DEFAULT '',
                ai_verified INTEGER NOT NULL DEFAULT 0,
                guven_puani INTEGER NOT NULL DEFAULT 1
            );
            CREATE UNIQUE INDEX IF NOT EXISTS idx_ana_havuz_company_norm ON ana_firma_havuzu(company_norm);
            """
        )

        # Mevcut tablolara güvenli sütun göçü (ALTER TABLE)
        _ensure_column(conn, "ana_firma_havuzu", "ai_verified", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(conn, "ana_firma_havuzu", "guven_puani", "INTEGER NOT NULL DEFAULT 1")
        _ensure_column(conn, "ana_firma_havuzu", "sektor", "TEXT NOT NULL DEFAULT ''")
        _ensure_column(conn, "mkys_gecmis_alimlar", "ai_verified", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(conn, "mkys_gecmis_alimlar", "guven_puani", "INTEGER NOT NULL DEFAULT 1")
        _ensure_column(conn, "mkys_gecmis_alimlar", "sektor", "TEXT NOT NULL DEFAULT ''")

        # Eski master_suppliers / mkys_history tabloları varsa verileri ai_verified=0 ile taşı
        tables = {r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
        if "master_suppliers" in tables:
            conn.execute(
                """
                INSERT OR IGNORE INTO ana_firma_havuzu(
                    company_name, company_norm, domain, website, primary_email,
                    all_emails_json, phone, sektor, items_sold_json, source, last_checked,
                    ai_verified, guven_puani
                )
                SELECT company_name, company_norm, domain, website, primary_email,
                       all_emails_json, phone, sector, items_sold_json, source, last_checked, 0, 1
                FROM master_suppliers
                """
            )
        if "mkys_history" in tables:
            conn.execute(
                """
                INSERT OR IGNORE INTO mkys_gecmis_alimlar(
                    item_name, item_norm, supplier_name, supplier_norm,
                    sektor, unit, quantity, unit_price, purchase_date,
                    ai_verified, guven_puani, created_at
                )
                SELECT item_name, item_norm, supplier_name, supplier_norm,
                       sector, unit, quantity, unit_price, purchase_date, 0, 1, created_at
                FROM mkys_history
                """
            )


def _touch_case(conn: sqlite3.Connection, case_id: int) -> None:
    conn.execute("UPDATE cases SET updated_at = ? WHERE id = ?", (_now_iso(), case_id))


# ---------------------------------------------------------------------------
# Varsayılan Komisyon & Ayarlar
# ---------------------------------------------------------------------------

def get_default_commission(db_path: Path | None = None) -> dict[str, Any]:
    with get_conn(db_path) as conn:
        row = conn.execute("SELECT value FROM app_settings WHERE key = 'default_commission'").fetchone()
        if not row:
            return dict(DEFAULT_COMMISSION)
        try:
            parsed = json.loads(row["value"])
            if isinstance(parsed, dict):
                members = parsed.get("members") or DEFAULT_COMMISSION["members"]
                return {
                    "members": members[:3] if members else DEFAULT_COMMISSION["members"][:1],
                    "authorizer_name": str(parsed.get("authorizer_name", "")),
                    "authorizer_title": str(parsed.get("authorizer_title", DEFAULT_COMMISSION["authorizer_title"])),
                }
        except Exception:
            pass
        return dict(DEFAULT_COMMISSION)


def save_default_commission(
    members: list[dict[str, str]],
    authorizer_name: str,
    authorizer_title: str,
    db_path: Path | None = None,
) -> None:
    payload = {
        "members": [{"name": m.get("name", "").strip(), "title": m.get("title", "").strip()} for m in members[:3]],
        "authorizer_name": authorizer_name.strip(),
        "authorizer_title": authorizer_title.strip(),
    }
    with get_conn(db_path) as conn:
        conn.execute(
            """
            INSERT INTO app_settings(key, value) VALUES('default_commission', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (json.dumps(payload, ensure_ascii=False),),
        )


# ---------------------------------------------------------------------------
# Alım Dosyaları CRUD
# ---------------------------------------------------------------------------

def list_cases(db_path: Path | None = None) -> list[dict[str, Any]]:
    with get_conn(db_path) as conn:
        rows = conn.execute(
            """
            SELECT c.*,
                (SELECT COUNT(*) FROM case_items ci WHERE ci.case_id = c.id) AS item_count,
                (SELECT COUNT(*) FROM case_suppliers cs WHERE cs.case_id = c.id) AS supplier_count,
                (SELECT COUNT(*) FROM case_offers co WHERE co.case_id = c.id) AS offer_count
            FROM cases c
            ORDER BY c.updated_at DESC, c.id DESC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def create_case(
    title: str = "Yeni Alım Dosyası",
    procurement_type: str = "Malzeme",
    db_path: Path | None = None,
) -> int:
    now = _now_iso()
    defaults = get_default_commission(db_path)
    with get_conn(db_path) as conn:
        cur = conn.execute(
            """
            INSERT INTO cases(
                title, procurement_type, deadline_date, deadline_time,
                evaluation_mode, commission_json, authorizer_name, authorizer_title,
                created_at, updated_at
            ) VALUES (?, ?, '', '', 'Kalem Bazlı En Düşük', ?, ?, ?, ?, ?)
            """,
            (
                title.strip() or "Yeni Alım Dosyası",
                procurement_type if procurement_type in ("Malzeme", "Hizmet") else "Malzeme",
                json.dumps(defaults["members"], ensure_ascii=False),
                defaults["authorizer_name"],
                defaults["authorizer_title"],
                now,
                now,
            ),
        )
        return int(cur.lastrowid)


def get_case(case_id: int, db_path: Path | None = None) -> dict[str, Any] | None:
    with get_conn(db_path) as conn:
        row = conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
        if not row:
            return None
        d = dict(row)
        try:
            d["commission"] = json.loads(d.get("commission_json") or "[]")
        except Exception:
            d["commission"] = []
        if not d["commission"]:
            d["commission"] = list(DEFAULT_COMMISSION["members"])
        return d


def update_case_meta(case_id: int, db_path: Path | None = None, **fields: Any) -> None:
    allowed = {
        "title", "procurement_type", "deadline_date", "deadline_time",
        "evaluation_mode", "commission_json", "authorizer_name", "authorizer_title",
    }
    updates: dict[str, Any] = {}
    for k, v in fields.items():
        if k == "commission" and isinstance(v, list):
            updates["commission_json"] = json.dumps(v[:3], ensure_ascii=False)
        elif k in allowed:
            updates[k] = v
    if not updates:
        return
    updates["updated_at"] = _now_iso()
    set_clause = ", ".join(f"{k} = ?" for k in updates)
    values = list(updates.values()) + [case_id]
    with get_conn(db_path) as conn:
        conn.execute(f"UPDATE cases SET {set_clause} WHERE id = ?", values)


def delete_case(case_id: int, db_path: Path | None = None) -> None:
    with get_conn(db_path) as conn:
        conn.execute("DELETE FROM cases WHERE id = ?", (case_id,))


# ---------------------------------------------------------------------------
# Kalemler CRUD
# ---------------------------------------------------------------------------

def get_case_items(case_id: int, db_path: Path | None = None) -> list[dict[str, Any]]:
    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT * FROM case_items WHERE case_id = ? ORDER BY sira_no ASC, id ASC",
            (case_id,),
        ).fetchall()
        out: list[dict[str, Any]] = []
        for r in rows:
            d = dict(r)
            try:
                d["ai_intel"] = json.loads(d.get("ai_intel_json") or "{}")
            except Exception:
                d["ai_intel"] = {}
            out.append(d)
        return out


def sync_case_items(case_id: int, items: list[dict[str, Any]], db_path: Path | None = None) -> list[dict[str, Any]]:
    """Mevcut kalem id'lerini koruyarak senkronize eder (teklif fiyat bağlarını bozmaz)."""
    with get_conn(db_path) as conn:
        existing_rows = conn.execute(
            "SELECT id FROM case_items WHERE case_id = ?", (case_id,)
        ).fetchall()
        existing_ids = {int(r["id"]) for r in existing_rows}
        kept_ids: set[int] = set()

        for idx, raw in enumerate(items, start=1):
            name = str(raw.get("name") or "").strip()
            if not name:
                continue
            try:
                sira_no = int(raw.get("sira_no") or idx)
            except (TypeError, ValueError):
                sira_no = idx
            unit = str(raw.get("unit") or "Adet").strip() or "Adet"
            try:
                qty = float(str(raw.get("quantity") or 1).replace(",", "."))
                if qty <= 0:
                    qty = 1.0
            except (TypeError, ValueError):
                qty = 1.0
            intel_json = json.dumps(raw.get("ai_intel") or {}, ensure_ascii=False) if raw.get("ai_intel") else None

            raw_id = raw.get("id")
            item_id: int | None = None
            if raw_id is not None and str(raw_id).strip() != "":
                try:
                    cand = int(float(raw_id))
                    if cand in existing_ids:
                        item_id = cand
                except (TypeError, ValueError):
                    pass

            if item_id is not None:
                kept_ids.add(item_id)
                if intel_json is not None:
                    conn.execute(
                        "UPDATE case_items SET sira_no=?, name=?, unit=?, quantity=?, ai_intel_json=? WHERE id=? AND case_id=?",
                        (sira_no, name, unit, qty, intel_json, item_id, case_id),
                    )
                else:
                    conn.execute(
                        "UPDATE case_items SET sira_no=?, name=?, unit=?, quantity=? WHERE id=? AND case_id=?",
                        (sira_no, name, unit, qty, item_id, case_id),
                    )
            else:
                conn.execute(
                    "INSERT INTO case_items(case_id, sira_no, name, unit, quantity, ai_intel_json) VALUES (?, ?, ?, ?, ?, ?)",
                    (case_id, sira_no, name, unit, qty, intel_json or "{}"),
                )

        to_delete = existing_ids - kept_ids
        if to_delete:
            placeholders = ",".join("?" for _ in to_delete)
            conn.execute(
                f"DELETE FROM case_items WHERE case_id = ? AND id IN ({placeholders})",
                [case_id, *to_delete],
            )
        _touch_case(conn, case_id)
    return get_case_items(case_id, db_path)


def save_item_intel(item_id: int, intel: dict[str, Any], db_path: Path | None = None) -> None:
    with get_conn(db_path) as conn:
        conn.execute(
            "UPDATE case_items SET ai_intel_json = ? WHERE id = ?",
            (json.dumps(intel, ensure_ascii=False), item_id),
        )


# ---------------------------------------------------------------------------
# Tedarikçiler CRUD
# ---------------------------------------------------------------------------

def get_case_suppliers(case_id: int, db_path: Path | None = None) -> list[dict[str, Any]]:
    with get_conn(db_path) as conn:
        rows = conn.execute(
            "SELECT * FROM case_suppliers WHERE case_id = ? ORDER BY item_name ASC, (CASE WHEN primary_email != '' THEN 0 ELSE 1 END), company_name ASC",
            (case_id,),
        ).fetchall()
        out: list[dict[str, Any]] = []
        for r in rows:
            d = dict(r)
            try:
                d["all_emails"] = json.loads(d.get("all_emails_json") or "[]")
            except Exception:
                d["all_emails"] = []
            d["selected"] = bool(d.get("selected", 1))
            out.append(d)
        return out


def add_case_suppliers(case_id: int, rows: list[dict[str, Any]], db_path: Path | None = None) -> int:
    """Bulunan firmaları aktif dosyaya mükerrer kontrolüyle ekler; eklenen kayıt sayısını döner."""
    added = 0
    with get_conn(db_path) as conn:
        existing = conn.execute(
            "SELECT id, item_name, domain, company_name, primary_email, all_emails_json FROM case_suppliers WHERE case_id = ?",
            (case_id,),
        ).fetchall()
        seen: dict[tuple[str, str], dict[str, Any]] = {}
        for ex in existing:
            key_dom = (ex["domain"] or ex["company_name"]).strip().lower()
            seen[(ex["item_name"].strip().lower(), key_dom)] = dict(ex)

        for r in rows:
            comp = str(r.get("company_name") or "").strip()
            if not comp:
                continue
            item_name = str(r.get("item") or r.get("item_name") or "").strip()
            dom = str(r.get("domain") or "").strip().lower()
            prim = str(r.get("primary_email") or "").strip().lower()
            all_em = [str(e).strip().lower() for e in (r.get("all_emails") or []) if str(e).strip()]
            if prim and prim not in all_em:
                all_em.insert(0, prim)
            key = (item_name.lower(), (dom or comp).lower())
            if key in seen:
                ex = seen[key]
                try:
                    old_em = json.loads(ex["all_emails_json"] or "[]")
                except Exception:
                    old_em = []
                merged_em = list(dict.fromkeys(old_em + all_em))
                new_prim = ex["primary_email"] or prim
                if ex["id"] > 0:
                    conn.execute(
                        "UPDATE case_suppliers SET primary_email=?, all_emails_json=? WHERE id=?",
                        (new_prim, json.dumps(merged_em, ensure_ascii=False), ex["id"]),
                    )
                continue

            selected = 1 if prim else 0
            conn.execute(
                """
                INSERT INTO case_suppliers(
                    case_id, item_name, company_name, domain, website,
                    primary_email, all_emails_json, phone, sector,
                    match_type, source, status, notes, selected
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    case_id,
                    item_name,
                    comp,
                    dom,
                    str(r.get("website") or ""),
                    prim,
                    json.dumps(all_em, ensure_ascii=False),
                    str(r.get("phone") or ""),
                    str(r.get("sector") or r.get("product_groups") or ""),
                    str(r.get("match_type") or ""),
                    str(r.get("source") or ""),
                    str(r.get("status") or "Aktif"),
                    str(r.get("notes") or ""),
                    selected,
                ),
            )
            seen[key] = {"id": -1, "primary_email": prim, "all_emails_json": "[]"}
            added += 1
        _touch_case(conn, case_id)
    return added


def update_supplier_selections_and_edits(
    case_id: int,
    edited_rows: list[dict[str, Any]],
    db_path: Path | None = None,
) -> None:
    with get_conn(db_path) as conn:
        for r in edited_rows:
            sid = r.get("id")
            if sid is None:
                continue
            prim = str(r.get("primary_email") or "").strip().lower()
            conn.execute(
                """
                UPDATE case_suppliers
                SET selected=?, company_name=?, primary_email=?, phone=?, website=?, item_name=?, sector=?
                WHERE id=? AND case_id=?
                """,
                (
                    1 if r.get("selected") else 0,
                    str(r.get("company_name") or "").strip(),
                    prim,
                    str(r.get("phone") or "").strip(),
                    str(r.get("website") or "").strip(),
                    str(r.get("item_name") or "").strip(),
                    str(r.get("sector") or "").strip(),
                    int(sid),
                    case_id,
                ),
            )
        _touch_case(conn, case_id)


def delete_case_suppliers(case_id: int, supplier_ids: list[int], db_path: Path | None = None) -> None:
    if not supplier_ids:
        return
    with get_conn(db_path) as conn:
        placeholders = ",".join("?" for _ in supplier_ids)
        conn.execute(
            f"DELETE FROM case_suppliers WHERE case_id = ? AND id IN ({placeholders})",
            [case_id, *[int(x) for x in supplier_ids]],
        )
        _touch_case(conn, case_id)


# ---------------------------------------------------------------------------
# Teklifler CRUD ve Uhde Değerlendirme
# ---------------------------------------------------------------------------

def get_case_offers(case_id: int, db_path: Path | None = None) -> list[dict[str, Any]]:
    with get_conn(db_path) as conn:
        offers = conn.execute(
            "SELECT * FROM case_offers WHERE case_id = ? ORDER BY id ASC", (case_id,)
        ).fetchall()
        prices = conn.execute(
            """
            SELECT p.offer_id, p.item_id, p.unit_price
            FROM case_offer_prices p
            JOIN case_offers o ON o.id = p.offer_id
            WHERE o.case_id = ?
            """,
            (case_id,),
        ).fetchall()
        price_map: dict[int, dict[int, float | None]] = {}
        for p in prices:
            price_map.setdefault(int(p["offer_id"]), {})[int(p["item_id"])] = p["unit_price"]
        out: list[dict[str, Any]] = []
        for o in offers:
            d = dict(o)
            d["prices"] = price_map.get(int(o["id"]), {})
            out.append(d)
        return out


def upsert_case_offer(
    case_id: int,
    firm_name: str,
    prices_by_item_id: dict[int, float | None],
    source_type: str = "Manuel",
    source_file: str = "",
    offer_id: int | None = None,
    db_path: Path | None = None,
) -> int:
    clean_firm = firm_name.strip() or "İsimsiz Firma"
    now = _now_iso()
    with get_conn(db_path) as conn:
        if offer_id is None:
            existing = conn.execute(
                "SELECT id FROM case_offers WHERE case_id = ? AND LOWER(firm_name) = LOWER(?)",
                (case_id, clean_firm),
            ).fetchone()
            if existing:
                offer_id = int(existing["id"])

        if offer_id is not None:
            conn.execute(
                "UPDATE case_offers SET firm_name=?, source_type=?, source_file=? WHERE id=? AND case_id=?",
                (clean_firm, source_type, source_file, offer_id, case_id),
            )
            conn.execute("DELETE FROM case_offer_prices WHERE offer_id = ?", (offer_id,))
        else:
            cur = conn.execute(
                "INSERT INTO case_offers(case_id, firm_name, source_type, source_file, created_at) VALUES (?, ?, ?, ?, ?)",
                (case_id, clean_firm, source_type, source_file, now),
            )
            offer_id = int(cur.lastrowid)

        for item_id, price in prices_by_item_id.items():
            val = float(price) if (price is not None and float(price) > 0) else None
            conn.execute(
                "INSERT INTO case_offer_prices(offer_id, item_id, unit_price) VALUES (?, ?, ?)",
                (offer_id, int(item_id), val),
            )
        _touch_case(conn, case_id)
        return offer_id


def delete_case_offer(case_id: int, offer_id: int, db_path: Path | None = None) -> None:
    with get_conn(db_path) as conn:
        conn.execute("DELETE FROM case_offers WHERE id = ? AND case_id = ?", (offer_id, case_id))
        _touch_case(conn, case_id)


def evaluate_case_offers(
    items: list[dict[str, Any]],
    offers: list[dict[str, Any]],
    evaluation_mode: str = "Kalem Bazlı En Düşük",
) -> dict[str, Any]:
    """Kalem bazlı veya toplam bazlı en uygun teklifi (uhdeyi) hesaplar."""
    item_results: list[dict[str, Any]] = []
    firm_totals: dict[int, dict[str, Any]] = {}

    for off in offers:
        oid = int(off["id"])
        quoted_count = 0
        total_amount = 0.0
        for it in items:
            iid = int(it["id"])
            qty = float(it.get("quantity") or 1.0)
            up = off.get("prices", {}).get(iid)
            if up is not None and up > 0:
                quoted_count += 1
                total_amount += up * qty
        firm_totals[oid] = {
            "offer_id": oid,
            "firm_name": off["firm_name"],
            "quoted_count": quoted_count,
            "is_complete": quoted_count == len(items) and len(items) > 0,
            "grand_total": round(total_amount, 2),
        }

    total_winner_id: int | None = None
    if evaluation_mode == "Toplam Bazlı En Düşük" and offers and items:
        complete = [ft for ft in firm_totals.values() if ft["is_complete"] and ft["grand_total"] > 0]
        pool = complete or [ft for ft in firm_totals.values() if ft["quoted_count"] > 0 and ft["grand_total"] > 0]
        if pool:
            best = min(pool, key=lambda x: (-x["quoted_count"], x["grand_total"]))
            total_winner_id = int(best["offer_id"])

    awarded_by_firm: dict[str, dict[str, Any]] = {}
    grand_awarded_total = 0.0

    for it in items:
        iid = int(it["id"])
        qty = float(it.get("quantity") or 1.0)
        firm_prices: dict[int, dict[str, Any]] = {}
        valid_candidates: list[tuple[float, int, str]] = []

        for off in offers:
            oid = int(off["id"])
            up = off.get("prices", {}).get(iid)
            line_tot = round(up * qty, 2) if (up is not None and up > 0) else None
            firm_prices[oid] = {"unit_price": up, "line_total": line_tot}
            if up is not None and up > 0:
                valid_candidates.append((float(up), oid, off["firm_name"]))

        winner_offer_id: int | None = None
        winner_firm = ""
        winner_unit: float | None = None
        winner_total: float | None = None

        if evaluation_mode == "Toplam Bazlı En Düşük":
            if total_winner_id is not None and firm_prices.get(total_winner_id, {}).get("unit_price"):
                winner_offer_id = total_winner_id
                winner_firm = firm_totals[total_winner_id]["firm_name"]
                winner_unit = float(firm_prices[total_winner_id]["unit_price"])
                winner_total = round(winner_unit * qty, 2)
        else:
            if valid_candidates:
                best_up, best_oid, best_firm = min(valid_candidates, key=lambda x: x[0])
                winner_offer_id = best_oid
                winner_firm = best_firm
                winner_unit = best_up
                winner_total = round(best_up * qty, 2)

        if winner_firm and winner_total is not None:
            grand_awarded_total += winner_total
            rec = awarded_by_firm.setdefault(winner_firm, {"firm_name": winner_firm, "item_count": 0, "total_amount": 0.0})
            rec["item_count"] += 1
            rec["total_amount"] = round(rec["total_amount"] + winner_total, 2)

        item_results.append({
            "item": it,
            "firm_prices": firm_prices,
            "winner_offer_id": winner_offer_id,
            "winner_firm": winner_firm,
            "winner_unit_price": winner_unit,
            "winner_total_price": winner_total,
        })

    return {
        "evaluation_mode": evaluation_mode,
        "item_results": item_results,
        "firm_totals": firm_totals,
        "total_winner_id": total_winner_id,
        "awarded_summary": sorted(awarded_by_firm.values(), key=lambda x: -x["total_amount"]),
        "grand_awarded_total": round(grand_awarded_total, 2),
    }


# ---------------------------------------------------------------------------
# MKYS Geçmiş Alım Kayıtları ve Sektör Etiketli Ana Firma Havuzu
# ---------------------------------------------------------------------------

def _fold_tr_text(text: str) -> str:
    low = str(text or "").replace("İ", "i").replace("I", "ı").lower()
    return low.translate(str.maketrans("çğıöşü", "cgiosu"))


def _norm_key(text: str) -> str:
    import re
    return re.sub(r"\s+", " ", _fold_tr_text(text)).strip()


def _extract_search_stems(text: str) -> set[str]:
    import re
    stop = {
        "ve", "veya", "ile", "icin", "grubu", "tipi", "adet", "kutu", "set", "paket",
        "malzeme", "malzemesi", "malzemeleri", "cihaz", "cihazi", "urun", "urunu",
        "satis", "toptan", "tedarik", "firma", "hizmet", "hizmeti", "alimi", "genel",
    }
    stems: set[str] = set()
    for w in re.split(r"[\s,/()\-&]+", _fold_tr_text(text)):
        clean = re.sub(r"[^a-z0-9]", "", w)
        if len(clean) < 3 or clean in stop:
            continue
        stems.add(clean)
        for sfx in ("lari", "leri", "lar", "ler", "nin", "nun", "si", "su", "i", "u", "e", "a"):
            if clean.endswith(sfx) and len(clean) - len(sfx) >= 4:
                stems.add(clean[: -len(sfx)])
                break
    return stems


def get_mkys_stats(db_path: Path | None = None) -> dict[str, int]:
    with get_conn(db_path) as conn:
        hc = conn.execute("SELECT COUNT(*) FROM mkys_gecmis_alimlar").fetchone()[0]
        sc = conn.execute("SELECT COUNT(DISTINCT supplier_norm) FROM mkys_gecmis_alimlar").fetchone()[0]
        mc = conn.execute("SELECT COUNT(*) FROM ana_firma_havuzu").fetchone()[0]
        vc = conn.execute("SELECT COUNT(*) FROM ana_firma_havuzu WHERE ai_verified = 1").fetchone()[0]
        return {
            "history_count": int(hc),
            "unique_suppliers": int(sc),
            "master_count": int(mc),
            "verified_count": int(vc),
        }


def import_mkys_records(
    records: list[dict[str, Any]],
    ai_verified: int = 0,
    db_path: Path | None = None,
) -> tuple[int, int]:
    """MKYS geçmiş alım satırlarını mkys_gecmis_alimlar ve ana_firma_havuzu tablolarına kaydeder."""
    now = _now_iso()
    inserted = 0
    with get_conn(db_path) as conn:
        for r in records:
            item_name = str(r.get("item_name") or "").strip()
            supplier_name = str(r.get("supplier_name") or "").strip()
            if not item_name or not supplier_name:
                continue
            item_norm = _norm_key(item_name)
            sup_norm = _norm_key(supplier_name)
            sector = str(r.get("sector") or "").strip()
            unit = str(r.get("unit") or "").strip()
            qty = r.get("quantity")
            up = r.get("unit_price")
            pdate = str(r.get("purchase_date") or "").strip()
            row_verified = int(r.get("ai_verified", ai_verified))

            cur = conn.execute(
                """
                INSERT INTO mkys_gecmis_alimlar(
                    item_name, item_norm, supplier_name, supplier_norm,
                    sektor, unit, quantity, unit_price, purchase_date,
                    ai_verified, guven_puani, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
                ON CONFLICT(item_norm, supplier_norm) DO UPDATE SET
                    sektor = CASE WHEN excluded.sektor != '' THEN excluded.sektor ELSE mkys_gecmis_alimlar.sektor END,
                    unit = CASE WHEN excluded.unit != '' THEN excluded.unit ELSE mkys_gecmis_alimlar.unit END,
                    quantity = COALESCE(excluded.quantity, mkys_gecmis_alimlar.quantity),
                    unit_price = COALESCE(excluded.unit_price, mkys_gecmis_alimlar.unit_price),
                    purchase_date = CASE WHEN excluded.purchase_date != '' THEN excluded.purchase_date ELSE mkys_gecmis_alimlar.purchase_date END,
                    ai_verified = MAX(mkys_gecmis_alimlar.ai_verified, excluded.ai_verified)
                """,
                (item_name, item_norm, supplier_name, sup_norm, sector, unit, qty, up, pdate, row_verified, now),
            )
            if cur.rowcount == 1:
                inserted += 1

            ex = conn.execute(
                "SELECT id, sektor, items_sold_json, ai_verified FROM ana_firma_havuzu WHERE company_norm = ?",
                (sup_norm,),
            ).fetchone()
            if ex:
                try:
                    sold = json.loads(ex["items_sold_json"] or "[]")
                except Exception:
                    sold = []
                if item_name not in sold:
                    sold = (sold + [item_name])[:50]
                old_sec = [s.strip() for s in str(ex["sektor"] or "").split(",") if s.strip()]
                if sector and sector.lower() not in {x.lower() for x in old_sec}:
                    old_sec.append(sector)
                conn.execute(
                    "UPDATE ana_firma_havuzu SET sektor=?, items_sold_json=?, ai_verified=? WHERE id=?",
                    (", ".join(old_sec[:6]), json.dumps(sold, ensure_ascii=False), max(int(ex["ai_verified"]), row_verified), ex["id"]),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO ana_firma_havuzu(
                        company_name, company_norm, domain, website, primary_email,
                        all_emails_json, phone, sektor, items_sold_json, source, last_checked,
                        ai_verified, guven_puani
                    ) VALUES (?, ?, '', '', '', '[]', '', ?, ?, 'MKYS Geçmiş Alım', '', ?, 1)
                    """,
                    (supplier_name, sup_norm, sector, json.dumps([item_name], ensure_ascii=False), row_verified),
                )

        total = int(conn.execute("SELECT COUNT(*) FROM mkys_gecmis_alimlar").fetchone()[0])
    return inserted, total


def upsert_master_supplier(
    info: dict[str, Any],
    sector: str = "",
    item_name: str = "",
    ai_verified: int = 0,
    guven_puani_delta: int = 0,
    db_path: Path | None = None,
) -> None:
    """Doğrulanmış veya havuz adayı firmayı ana_firma_havuzu tablosuna kaydeder."""
    comp = str(info.get("company_name") or "").strip()
    if not comp:
        return
    comp_norm = _norm_key(comp)
    dom = str(info.get("domain") or "").strip().lower()
    prim = str(info.get("primary_email") or "").strip().lower()
    all_em = [str(e).strip().lower() for e in (info.get("all_emails") or []) if str(e).strip()]
    if prim and prim not in all_em:
        all_em.insert(0, prim)
    phone = str(info.get("phone") or "").strip()
    website = str(info.get("website") or "").strip()
    sec_in = str(sector or info.get("sector") or info.get("sektor") or "").strip()
    verified_flag = int(info.get("ai_verified", ai_verified))

    with get_conn(db_path) as conn:
        ex = None
        if dom:
            ex = conn.execute("SELECT * FROM ana_firma_havuzu WHERE LOWER(domain) = ?", (dom,)).fetchone()
        if ex is None:
            ex = conn.execute("SELECT * FROM ana_firma_havuzu WHERE company_norm = ?", (comp_norm,)).fetchone()

        if ex:
            try:
                old_em = json.loads(ex["all_emails_json"] or "[]")
            except Exception:
                old_em = []
            merged_em = list(dict.fromkeys(old_em + all_em))
            try:
                sold = json.loads(ex["items_sold_json"] or "[]")
            except Exception:
                sold = []
            if item_name and item_name not in sold:
                sold = (sold + [item_name])[:50]
            # Gemini doğruladıysa gerçek sektörü başa/güncellemeye yaz
            if verified_flag and sec_in:
                new_sector = sec_in
            else:
                old_sec = [s.strip() for s in str(ex["sektor"] or "").split(",") if s.strip()]
                for piece in [x.strip() for x in sec_in.split(",") if x.strip()]:
                    if piece.lower() not in {x.lower() for x in old_sec}:
                        old_sec.append(piece)
                new_sector = ", ".join(old_sec[:6])

            new_trust = max(1, int(ex["guven_puani"] or 1) + int(guven_puani_delta))
            new_verified = max(int(ex["ai_verified"] or 0), verified_flag)
            conn.execute(
                """
                UPDATE ana_firma_havuzu
                SET company_name=?, domain=?, website=?, primary_email=?, all_emails_json=?,
                    phone=?, sektor=?, items_sold_json=?, source=?, last_checked=?,
                    ai_verified=?, guven_puani=?
                WHERE id=?
                """,
                (
                    comp if len(comp) >= 3 else ex["company_name"],
                    dom or ex["domain"],
                    website or ex["website"],
                    prim or ex["primary_email"],
                    json.dumps(merged_em, ensure_ascii=False),
                    phone or ex["phone"],
                    new_sector,
                    json.dumps(sold, ensure_ascii=False),
                    str(info.get("source") or ex["source"] or "Ana Havuz"),
                    dt.date.today().isoformat(),
                    new_verified,
                    new_trust,
                    ex["id"],
                ),
            )
        else:
            sold = [item_name] if item_name else []
            conn.execute(
                """
                INSERT INTO ana_firma_havuzu(
                    company_name, company_norm, domain, website, primary_email,
                    all_emails_json, phone, sektor, items_sold_json, source, last_checked,
                    ai_verified, guven_puani
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    comp, comp_norm, dom, website, prim,
                    json.dumps(all_em, ensure_ascii=False),
                    phone, sec_in,
                    json.dumps(sold, ensure_ascii=False),
                    str(info.get("source") or "Ana Havuz"),
                    dt.date.today().isoformat(),
                    verified_flag,
                    max(1, 1 + int(guven_puani_delta)),
                ),
            )


def seed_unverified_from_xlsx(cache_rows: dict[str, dict[str, Any]], db_path: Path | None = None) -> int:
    """Eski firma_havuzu.xlsx kayıtlarını henüz ana_firma_havuzu'nda yoksa ai_verified=0 olarak aktarır."""
    if not cache_rows:
        return 0
    added = 0
    with get_conn(db_path) as conn:
        for dom, entry in cache_rows.items():
            comp = str(entry.get("company_name") or dom).strip()
            comp_norm = _norm_key(comp)
            ex = conn.execute(
                "SELECT id FROM ana_firma_havuzu WHERE LOWER(domain) = ? OR company_norm = ?",
                (dom.lower(), comp_norm),
            ).fetchone()
            if ex:
                continue
            conn.execute(
                """
                INSERT OR IGNORE INTO ana_firma_havuzu(
                    company_name, company_norm, domain, website, primary_email,
                    all_emails_json, phone, sektor, items_sold_json, source, last_checked,
                    ai_verified, guven_puani
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '[]', 'Excel Havuz Göçü', ?, 0, 1)
                """,
                (
                    comp,
                    comp_norm,
                    dom.lower(),
                    str(entry.get("website") or ""),
                    str(entry.get("primary_email") or "").lower(),
                    json.dumps(entry.get("all_emails") or [], ensure_ascii=False),
                    str(entry.get("phone") or ""),
                    str(entry.get("sector") or entry.get("product_groups") or ""),
                    str(entry.get("last_checked") or ""),
                ),
            )
            added += 1
    return added


def get_unverified_pool_records(limit: int = 200, db_path: Path | None = None) -> dict[str, list[dict[str, Any]]]:
    """ai_verified = 0 olan ana_firma_havuzu ve mkys_gecmis_alimlar kayıtlarını döner."""
    with get_conn(db_path) as conn:
        pool_rows = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM ana_firma_havuzu WHERE ai_verified = 0 ORDER BY id ASC LIMIT ?",
                (limit,),
            ).fetchall()
        ]
        mkys_rows = [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM mkys_gecmis_alimlar WHERE ai_verified = 0 ORDER BY id ASC LIMIT ?",
                (limit,),
            ).fetchall()
        ]
    return {"pool": pool_rows, "mkys": mkys_rows}


def apply_pool_sanitization_results(
    pool_decisions: list[dict[str, Any]],
    mkys_decisions: list[dict[str, Any]] | None = None,
    db_path: Path | None = None,
) -> dict[str, int]:
    """Gemini havuz denetim kararlarına göre çöp kayıtları siler, gerçek firmaları ai_verified=1 yapar."""
    deleted = 0
    verified = 0
    with get_conn(db_path) as conn:
        for d in pool_decisions:
            rid = int(d["id"])
            if not d.get("keep"):
                conn.execute("DELETE FROM ana_firma_havuzu WHERE id = ?", (rid,))
                deleted += 1
            else:
                new_sec = str(d.get("real_sector") or "").strip()
                new_name = str(d.get("clean_name") or "").strip()
                if new_name:
                    # company_norm çakışması varsa mevcut kaydı koruyarak güncelle
                    new_norm = _norm_key(new_name)
                    conflict = conn.execute(
                        "SELECT id FROM ana_firma_havuzu WHERE company_norm = ? AND id != ?",
                        (new_norm, rid),
                    ).fetchone()
                    if conflict:
                        conn.execute("DELETE FROM ana_firma_havuzu WHERE id = ?", (rid,))
                        deleted += 1
                        continue
                    conn.execute(
                        "UPDATE ana_firma_havuzu SET company_name=?, company_norm=?, sektor=?, ai_verified=1 WHERE id=?",
                        (new_name, new_norm, new_sec, rid),
                    )
                else:
                    conn.execute(
                        "UPDATE ana_firma_havuzu SET sektor=?, ai_verified=1 WHERE id=?",
                        (new_sec, rid),
                    )
                verified += 1

        for md in (mkys_decisions or []):
            mid = int(md["id"])
            if not md.get("keep"):
                conn.execute("DELETE FROM mkys_gecmis_alimlar WHERE id = ?", (mid,))
                deleted += 1
            else:
                new_sec = str(md.get("real_sector") or "").strip()
                conn.execute(
                    "UPDATE mkys_gecmis_alimlar SET sektor=?, ai_verified=1 WHERE id=?",
                    (new_sec, mid),
                )
                verified += 1
    return {"deleted": deleted, "verified": verified}


def increment_supplier_trust(
    identifiers: list[str],
    delta: int = 1,
    db_path: Path | None = None,
) -> int:
    """Adım 3 davet veya Adım 4 teklif girişinde firmanın guven_puani değerini artırır."""
    cleaned = [str(x or "").strip() for x in identifiers if str(x or "").strip()]
    if not cleaned:
        return 0
    updated = 0
    with get_conn(db_path) as conn:
        for ident in cleaned:
            low = ident.lower()
            norm = _norm_key(ident)
            cur1 = conn.execute(
                "UPDATE ana_firma_havuzu SET guven_puani = COALESCE(guven_puani, 1) + ? WHERE LOWER(domain) = ? OR company_norm = ?",
                (delta, low, norm),
            )
            cur2 = conn.execute(
                "UPDATE mkys_gecmis_alimlar SET guven_puani = COALESCE(guven_puani, 1) + ? WHERE supplier_norm = ?",
                (delta, norm),
            )
            updated += int(cur1.rowcount or 0) + int(cur2.rowcount or 0)
    return updated


def query_mkys_and_master_suppliers(
    item_name: str,
    sector: str = "",
    synonyms: list[str] | None = None,
    db_path: Path | None = None,
) -> list[dict[str, Any]]:
    """Yalnızca ai_verified = 1 olan ve gerçek faaliyet alanı kalemle uyuşan havuz/MKYS firmalarını guven_puani sırasıyla döner."""
    item_stems = _extract_search_stems(" ".join([item_name] + list(synonyms or [])))
    sector_stems = _extract_search_stems(sector)
    if not item_stems and not sector_stems:
        return []

    candidates: dict[str, dict[str, Any]] = {}
    with get_conn(db_path) as conn:
        master_rows = conn.execute(
            "SELECT * FROM ana_firma_havuzu WHERE ai_verified = 1 ORDER BY guven_puani DESC, id ASC"
        ).fetchall()
        for mr in master_rows:
            d = dict(mr)
            try:
                sold = json.loads(d.get("items_sold_json") or "[]")
            except Exception:
                sold = []
            try:
                all_em = json.loads(d.get("all_emails_json") or "[]")
            except Exception:
                all_em = []
            sold_text = " ".join(str(x) for x in sold)
            sold_stems = _extract_search_stems(sold_text)
            sec_stems = _extract_search_stems(d.get("sektor", ""))

            item_hits = len(item_stems & (sold_stems | sec_stems))
            sec_hits = len(sector_stems & (sec_stems | sold_stems))
            score = item_hits * 4 + sec_hits * 2
            if score <= 0:
                continue

            cnorm = d["company_norm"]
            trust = int(d.get("guven_puani") or 1)
            candidates[cnorm] = {
                "score": score,
                "guven_puani": trust,
                "item_hits": item_hits,
                "company_name": d["company_name"],
                "company_norm": cnorm,
                "domain": d["domain"],
                "website": d["website"],
                "primary_email": d["primary_email"],
                "all_emails": all_em,
                "phone": d["phone"],
                "sector": d["sektor"] or sector,
                "matched_items": sold[:5],
                "source": "MKYS Geçmiş Alım" if "MKYS" in (d.get("source") or "") else "Ana Firma Havuzu (AI Doğrulamalı)",
                "match_type": "Havuz Kalem Eşleşmesi" if item_hits > 0 else "Havuz Sektör Eşleşmesi",
                "needs_contact_lookup": not bool(d["primary_email"] or d["phone"]),
                "ai_verified": 1,
            }

        mkys_rows = conn.execute(
            "SELECT * FROM mkys_gecmis_alimlar WHERE ai_verified = 1 ORDER BY guven_puani DESC, id ASC"
        ).fetchall()
        for kr in mkys_rows:
            knorm = kr["supplier_norm"]
            row_stems = _extract_search_stems(kr["item_name"])
            row_sec_stems = _extract_search_stems(kr["sektor"])
            item_hits = len(item_stems & row_stems)
            sec_hits = len(sector_stems & (row_sec_stems | row_stems))
            score = item_hits * 5 + sec_hits * 2
            if score <= 0:
                continue
            trust = int(kr["guven_puani"] or 1)
            if knorm in candidates:
                candidates[knorm]["score"] = max(candidates[knorm]["score"], score)
                candidates[knorm]["guven_puani"] = max(candidates[knorm]["guven_puani"], trust)
                if item_hits > 0:
                    candidates[knorm]["match_type"] = "MKYS Geçmiş Alım"
                    candidates[knorm]["source"] = (
                        "MKYS Geçmiş Alım (Kayıtlı İletişim)"
                        if not candidates[knorm]["needs_contact_lookup"]
                        else "MKYS Geçmiş Alım"
                    )
            else:
                candidates[knorm] = {
                    "score": score,
                    "guven_puani": trust,
                    "item_hits": item_hits,
                    "company_name": kr["supplier_name"],
                    "company_norm": knorm,
                    "domain": "",
                    "website": "",
                    "primary_email": "",
                    "all_emails": [],
                    "phone": "",
                    "sector": kr["sektor"] or sector,
                    "matched_items": [kr["item_name"]],
                    "source": "MKYS Geçmiş Alım",
                    "match_type": "MKYS Geçmiş Alım" if item_hits > 0 else "MKYS Sektör Eşleşmesi",
                    "needs_contact_lookup": True,
                    "ai_verified": 1,
                }

    ranked = sorted(
        candidates.values(),
        key=lambda x: (-x["guven_puani"], -x["score"], 0 if not x["needs_contact_lookup"] else 1, x["company_name"]),
    )
    return ranked

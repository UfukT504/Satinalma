"""Satın alma birimi için sıfır kredili masaüstü kontrol paneli.
Satın alma formundan yapıştırılan çok satırlı tabloları veya Word/Excel/PDF dosyalarını
doğrudan işler; canlı ilerleme çubuğu, sayaç ve anlık sonuç akışı sunar.
"""

from __future__ import annotations

import datetime as dt
import os
import re
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / ".agent" / "skills" / "satinalma-tedarik" / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import parse_form  # noqa: E402
import scrape_suppliers  # noqa: E402

OUTPUT_XLSX = BASE_DIR / "tedarikci_listesi.xlsx"
CACHE_XLSX = BASE_DIR / "firma_havuzu.xlsx"

IGNORE_HEADER_WORDS = {
    "SIRA NO",
    "MALZEME ADI",
    "MALZEME ADI VE KISACA AÇIKLAMASI",
    "MALZEME ADI VE KISACA AÇIKMLAMASI",
    "BİRİM",
    "MİKTAR",
    "ADET",
}


def extract_items_from_pasted_text(raw_text: str) -> list[str]:
    """Yapıştırılan tablo veya virgüllü metinden temiz kalem adlarını çıkarır."""
    items: list[str] = []
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
    for line in lines:
        if "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
            if len(parts) >= 2 and re.fullmatch(r"\d+[.)]?", parts[0]):
                candidate = parts[1]
            else:
                candidate = parts[0]
            if candidate.upper() in IGNORE_HEADER_WORDS or len(candidate) < 3:
                continue
            items.append(candidate)
        else:
            for chunk in line.split(","):
                c = re.sub(r"^\d+[.)]\s*", "", chunk.strip())
                if c and c.upper() not in IGNORE_HEADER_WORDS and len(c) >= 3:
                    items.append(c)
    return list(dict.fromkeys(items))


class TedarikPaneli(tk.Tk):
    """Canlı ilerleme çubuğu ve anlık sonuç akışlı Türkçe masaüstü arayüzü."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Satın Alma - Tedarikçi & E-posta Bulucu (0 Kredi)")
        self.geometry("1060x700")
        self._stop_event = threading.Event()
        self._is_running = False
        self._start_ts = 0.0
        self._rows_by_iid: dict[str, dict] = {}
        self._base_status = f"Hazır. Proje klasörü: {BASE_DIR}"
        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=12)
        top.pack(fill=tk.X)

        ttk.Label(
            top,
            text="Kalemler (Word/Excel/PDF seçebilir veya Excel tablosunu direkt yapıştırabilirsiniz):",
        ).pack(anchor=tk.W)

        row1 = ttk.Frame(top)
        row1.pack(fill=tk.X, pady=(4, 8))

        self.item_text = tk.Text(row1, height=4, font=("Segoe UI", 9))
        self.item_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        btn_col = ttk.Frame(row1)
        btn_col.pack(side=tk.LEFT, fill=tk.Y)
        self.search_btn = ttk.Button(btn_col, text="Tedarikçi & Mail Bul", command=self._on_search)
        self.search_btn.pack(fill=tk.X, pady=(0, 4))
        self.form_btn = ttk.Button(btn_col, text="Word / Excel / PDF Ekle...", command=self._on_pick_file)
        self.form_btn.pack(fill=tk.X, pady=(0, 4))
        self.mkys_btn = ttk.Button(btn_col, text="MKYS Firmalarını Tara...", command=self._on_scan_mkys)
        self.mkys_btn.pack(fill=tk.X, pady=(0, 4))
        self.stop_btn = ttk.Button(btn_col, text="Durdur", command=self._on_stop, state=tk.DISABLED)
        self.stop_btn.pack(fill=tk.X)

        ttk.Label(top, text="Ürün Grubu / Branş (opsiyonel - boş bırakırsanız otomatik tahmin edilir):").pack(anchor=tk.W)
        self.group_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.group_var, font=("Segoe UI", 9)).pack(fill=tk.X, pady=(2, 8))

        row2 = ttk.Frame(top)
        row2.pack(fill=tk.X, pady=(0, 4))
        self.status_var = tk.StringVar(value=self._base_status)
        ttk.Label(row2, textvariable=self.status_var, foreground="#1F4E78", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        ttk.Button(row2, text="Temizle", command=self._on_clear).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(row2, text="Klasörü Aç", command=lambda: os.startfile(str(BASE_DIR))).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(row2, text="Excel'i Aç", command=self._open_excel).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(row2, text="Mailleri Kopyala (; ile)", command=lambda: self._copy_emails("; ")).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(row2, text="Mailleri Kopyala (Alt Alta)", command=lambda: self._copy_emails("\n")).pack(side=tk.RIGHT)

        self.progress = ttk.Progressbar(top, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, pady=(2, 0))

        cols = ("Kalem", "Firma Adı", "Eşleşme Tipi", "Öncelikli E-posta", "Telefon", "Ürün Grupları", "Web Sitesi")
        table_frame = ttk.Frame(self, padding=(12, 0, 12, 12))
        table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="extended")
        for c, w in zip(cols, (150, 170, 150, 175, 120, 140, 170)):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=tk.W)

        yscroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

    def _on_clear(self) -> None:
        if self._is_running:
            return
        self.item_text.delete("1.0", tk.END)
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._rows_by_iid.clear()
        self._rows_by_iid.clear()
        self.progress["value"] = 0
        self._set_status("Temizlendi. Yeni kalem ekleyebilirsiniz.")

    def _copy_emails(self, separator: str) -> None:
        """Tablodaki seçili (veya tüm) geçerli e-postaları doğrudan mail alıcı kutusuna yapıştırılacak formatta kopyalar."""
        selected = self.tree.selection()
        target_iids = selected if selected else self.tree.get_children()
        emails: list[str] = []
        seen: set[str] = set()
        for iid in target_iids:
            row = self._rows_by_iid.get(iid) or {}
            candidates = list(row.get("all_emails") or [])
            if row.get("primary_email") and row["primary_email"] not in candidates:
                candidates.insert(0, row["primary_email"])
            for em in candidates:
                em_l = em.strip().lower()
                if not em_l or em_l == "-" or em_l.endswith(".kep.tr") or em_l in seen:
                    continue
                seen.add(em_l)
                emails.append(em_l)
        if not emails:
            messagebox.showinfo("Bilgi", "Kopyalanacak geçerli e-posta bulunamadı.")
            return
        self.clipboard_clear()
        self.clipboard_append(separator.join(emails))
        self.update()
        mode_name = "alt alta" if separator == "\n" else "'; ' ile ayrılmış"
        self._set_status(f"{len(emails)} adet e-posta panoya kopyalandı ({mode_name}). Mailde Kime/Gizli kısmına yapıştırabilirsiniz.")

    def _on_stop(self) -> None:
        self._stop_event.set()
        self.stop_btn.configure(state=tk.DISABLED)
        self._set_status("Durduruluyor, şimdiye kadar bulunanlar kaydediliyor...")

    def _set_status(self, msg: str) -> None:
        self._base_status = msg
        if self._is_running:
            elapsed = int(time.time() - self._start_ts)
            self.status_var.set(f"[{elapsed} sn] {msg}")
        else:
            self.status_var.set(msg)

    def _tick_timer(self) -> None:
        if not self._is_running:
            return
        elapsed = int(time.time() - self._start_ts)
        self.status_var.set(f"[{elapsed} sn] {self._base_status}")
        self.after(1000, self._tick_timer)

    def _open_excel(self) -> None:
        if OUTPUT_XLSX.exists():
            os.startfile(str(OUTPUT_XLSX))
        else:
            messagebox.showinfo("Bilgi", "Henüz tedarikci_listesi.xlsx oluşturulmadı.")

    def _on_pick_file(self) -> None:
        path_str = filedialog.askopenfilename(
            title="Word Talep Listesi, MKYS Excel veya PDF Seç",
            filetypes=[
                ("Word / Excel / PDF", "*.docx *.xlsx *.xlsm *.pdf *.csv *.txt"),
                ("Word Belgeleri (*.docx)", "*.docx"),
                ("Excel Dosyaları (*.xlsx)", "*.xlsx *.xlsm"),
                ("Tüm Dosyalar", "*.*"),
            ],
        )
        if not path_str:
            return
        path = Path(path_str)
        if path.suffix.lower() == ".doc":
            messagebox.showwarning(
                "Eski Word Formatı (.doc)",
                "Bu dosya eski .doc formatında. Lütfen Word'de 'Farklı Kaydet > Word Belgesi (.docx)' yapıp tekrar seçin.",
            )
            return
        readers = {
            ".xlsx": parse_form.read_xlsx,
            ".xlsm": parse_form.read_xlsx,
            ".csv": parse_form.read_csv,
            ".docx": parse_form.read_docx,
            ".pdf": parse_form.read_pdf,
            ".txt": parse_form.read_text,
        }
        reader = readers.get(path.suffix.lower())
        if not reader:
            messagebox.showerror("Hata", f"Desteklenmeyen dosya türü: {path.suffix}")
            return

        data = reader(path)
        items = [x["item"] for x in data.get("items", []) if x.get("item")]
        existing = extract_items_from_pasted_text(self.item_text.get("1.0", tk.END))
        self.item_text.delete("1.0", tk.END)
        if items:
            merged = list(dict.fromkeys(existing + items))
            self.item_text.insert("1.0", "\n".join(merged[:50]))
            stats = data.get("price_stats") or {}
            msg = f"Dosyadan {len(items)} kalem eklendi (toplam {len(merged)} kalem)."
            if data.get("companies"):
                msg += f" {len(data['companies'])} firma görüldü."
            if stats.get("avg"):
                msg += f" Genel ort. fiyat: {stats['avg']}"
            self._set_status(msg)
        else:
            raw = data.get("raw_lines") or data.get("paragraphs") or []
            if raw:
                merged = list(dict.fromkeys(existing + raw[:15]))
                self.item_text.insert("1.0", "\n".join(merged))
                self._set_status(f"Dosyadan {len(raw)} satır metin eklendi; gereksiz satırları silip aratabilirsiniz.")
            else:
                messagebox.showwarning("Uyarı", "Dosyada okunabilir kalem bulunamadı.")

    def _on_search(self) -> None:
        raw = self.item_text.get("1.0", tk.END).strip()
        items = extract_items_from_pasted_text(raw)
        if not items:
            messagebox.showwarning("Uyarı", "Lütfen en az bir kalem adı yazın veya Word/Excel ekleyin.")
            return
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._rows_by_iid.clear()
        group_queries = [f"{g.strip()} medikal tedarikçi" for g in self.group_var.get().split(",") if g.strip()]
        self._stop_event.clear()
        self._is_running = True
        self._start_ts = time.time()
        self.progress["value"] = 2
        self.search_btn.configure(state=tk.DISABLED)
        self.mkys_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self._set_status(f"Başlatılıyor ({len(items)} kalem — tahmini süre ~{len(items) * 20}-{len(items) * 35} sn)...")
        self.after(1000, self._tick_timer)
        threading.Thread(target=self._run_worker, args=(items, group_queries), daemon=True).start()

    def _on_scan_mkys(self) -> None:
        path_str = filedialog.askopenfilename(
            title="MKYS Fiyat Sorgu Excel/CSV Dosyası Seç",
            filetypes=[("MKYS Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Tüm Dosyalar", "*.*")],
        )
        if not path_str:
            return
        path = Path(path_str)
        data = parse_form.read_csv(path) if path.suffix.lower() == ".csv" else parse_form.read_xlsx(path)
        details = data.get("company_details") or []
        private_firms = [c for c in details if not c.get("is_public")]
        if not private_firms:
            messagebox.showwarning("Uyarı", "Bu MKYS dosyasında taranabilir özel tedarikçi firma adı bulunamadı.")
            return

        for i in self.tree.get_children():
            self.tree.delete(i)
        self._rows_by_iid.clear()
        item_label = data["items"][0]["item"] if data.get("items") else "MKYS Havuz"
        stats = data.get("price_stats") or {}
        avg_txt = f" | Ort. Alım Fiyatı: {stats['avg']}" if stats.get("avg") else ""

        self._stop_event.clear()
        self._is_running = True
        self._start_ts = time.time()
        self.progress["value"] = 5
        self.search_btn.configure(state=tk.DISABLED)
        self.mkys_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self._set_status(f"MKYS okundu: {len(private_firms)} firma taranacak ({item_label}{avg_txt})...")
        self.after(1000, self._tick_timer)
        threading.Thread(target=self._run_mkys_worker, args=(private_firms, item_label), daemon=True).start()

    def _run_mkys_worker(self, company_details: list[dict], item_label: str) -> None:
        cache = scrape_suppliers.load_cache(CACHE_XLSX)

        def on_progress(msg: str) -> None:
            self.after(0, lambda m=msg: self._set_status(m))

        rows = scrape_suppliers.resolve_mkys_companies(
            company_details, item_label, cache, progress_cb=on_progress, stop_event=self._stop_event
        )
        self._save_outputs(rows, cache)
        self.after(0, lambda: self._append_rows_ui(rows, 100.0))
        self.after(0, lambda: self._finish_ui(rows))

    def _save_outputs(self, all_rows: list[dict], cache: dict[str, dict]) -> None:
        scrape_suppliers.export_results(OUTPUT_XLSX, CACHE_XLSX, all_rows, cache)

    def _append_rows_ui(self, new_rows: list[dict], pct: float) -> None:
        for r in new_rows:
            iid = self.tree.insert(
                "",
                tk.END,
                values=(
                    r["item"],
                    r["company_name"],
                    r.get("match_type", "-"),
                    r["primary_email"] or "-",
                    r["phone"] or "-",
                    r.get("product_groups", "-"),
                    r["website"],
                ),
            )
            self._rows_by_iid[iid] = r
        self.progress["value"] = pct

    def _run_worker(self, items: list[str], group_queries: list[str]) -> None:
        cache = scrape_suppliers.load_cache(CACHE_XLSX)
        all_rows: list[dict] = []
        total = len(items)
        for idx, item in enumerate(items, start=1):
            if self._stop_event.is_set():
                break

            def on_progress(msg: str, i: int = idx) -> None:
                self.after(0, lambda m=f"[{i}/{total}] {msg}": self._set_status(m))

            smart = scrape_suppliers.build_smart_queries(item)
            task = {
                "item": item,
                "domains": [],
                "queries": smart["queries"],
                "group_queries": group_queries,
                "social_queries": smart["social_queries"],
                "auto_expand": True,
            }
            item_rows = scrape_suppliers.process_task(
                task, cache, progress_cb=on_progress, stop_event=self._stop_event
            )
            all_rows.extend(item_rows)
            self._save_outputs(all_rows, cache)
            pct = (idx / total) * 100.0
            self.after(0, lambda r=item_rows, p=pct: self._append_rows_ui(r, p))

        self.after(0, lambda: self._finish_ui(all_rows))

    def _finish_ui(self, rows: list[dict]) -> None:
        self._is_running = False
        elapsed = int(time.time() - self._start_ts)
        self.progress["value"] = 100 if not self._stop_event.is_set() else self.progress["value"]
        self.search_btn.configure(state=tk.NORMAL)
        self.mkys_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)
        prefix = "Durduruldu" if self._stop_event.is_set() else "Tamamlandı"
        self._set_status(
            f"{prefix} ({elapsed} sn): {len(rows)} aktif tedarikçi bulundu -> tedarikci_listesi.xlsx kaydedildi."
        )


if __name__ == "__main__":
    app = TedarikPaneli()
    app.mainloop()

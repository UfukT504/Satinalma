"""Resmi Word (.docx) Belge Üreticileri:
1) Adım 3: Resmi Teklife Davet Mektubu (Times New Roman 12pt, iki yana yaslı)
2) Adım 4: Piyasa Fiyat Araştırma Tutanağı (modüler, ileride özel şablon .docx bağlanabilir)
"""

from __future__ import annotations

import datetime as dt
import io
from pathlib import Path
from typing import Any


def format_tr_number(val: float | int | None, decimals: int = 2) -> str:
    if val is None:
        return "-"
    num = float(val)
    if decimals == 0 or (decimals == 2 and num.is_integer() and False):
        pass
    formatted = f"{num:,.{decimals}f}"
    # 1,234.56 -> 1.234,56
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def format_qty(val: float | int | None) -> str:
    if val is None or val == "":
        return ""
    num = float(val)
    if num.is_integer():
        return str(int(num))
    return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".").rstrip("0").rstrip(",")


def build_invitation_text(deadline_date: str = "", deadline_time: str = "") -> tuple[str, str]:
    """Kullanıcının belirttiği birebir resmi teklife davet metnini döndürür."""
    d_str = deadline_date.strip() or "..../..../2026"
    t_str = deadline_time.strip() or "...:..."
    p1 = (
        "Hastanemiz bünyesinde hizmet vermekte olan ilgili birimlerde aktif olarak kullanılan ve temin ihtiyacı bulunan, "
        "aşağıda belirtilen kalemlere ilişkin teknik şartname detaylarını incelemenizi; teklif vereceğiniz kalemlere ait "
        "fiyat teklifinizi hazırlayarak teknik şartname ve teklif mektubunu firma kaşesi ile yetkili kişi imzasını havi şekilde "
        f"{d_str} tarihi saat {t_str}'e kadar elden ve/veya ufuk.tonguc@saglik.gov.tr e-posta adresine ulaştırmanızı hususunda;"
    )
    p2 = "Gereğini bilgilerinize rica ederim."
    return p1, p2


def build_rfq_draft(items: list[str], deadline_date: str = "", deadline_time: str = "") -> tuple[str, str]:
    items_str = "\n".join(f"- {idx}. {it}" for idx, it in enumerate(items, 1))
    p1, p2 = build_invitation_text(deadline_date, deadline_time)
    subject = f"4734/22-d Doğrudan Temin Teklife Davet — {', '.join(items[:2])}" if items else "4734/22-d Doğrudan Temin Teklife Davet"
    body = f"Sayın Yetkili,\n\n{p1}\n{p2}\n\nKalemler:\n{items_str}\n\nSatın Alma Birimi"
    return subject, body


def _apply_tnr(run: Any, size_pt: int = 12, bold: bool = False, color_rgb: tuple[int, int, int] = (0, 0, 0)) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Times New Roman")


def _set_table_borders(table: Any, color: str = "000000", sz: str = "6") -> None:
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        "</w:tblBorders>"
    )
    tbl_pr.append(borders)


def _remove_table_borders(table: Any) -> None:
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    tbl_pr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        "</w:tblBorders>"
    )
    tbl_pr.append(borders)


def _format_cell(
    cell: Any,
    text: str,
    width: Any,
    align: Any,
    size_pt: int = 11,
    bold: bool = False,
    shade: str = "",
    color_rgb: tuple[int, int, int] = (0, 0, 0),
) -> None:
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Pt

    cell.width = width
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if shade:
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{shade}"/>'))
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    _apply_tnr(p.add_run(str(text)), size_pt=size_pt, bold=bold, color_rgb=color_rgb)


# ---------------------------------------------------------------------------
# Adım 3: Resmi Teklife Davet Mektubu (.docx)
# ---------------------------------------------------------------------------

def build_invitation_docx(
    table_rows: list[dict[str, Any]],
    procurement_type: str,
    deadline_date: str = "",
    deadline_time: str = "",
    case_title: str = "",
) -> bytes:
    """Times New Roman 12pt, iki yana yaslı resmi Teklife Davet Mektubu (.docx) üretir."""
    import docx
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Cm, Pt

    doc = docx.Document()
    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.2)
        sec.left_margin = Cm(2.0)
        sec.right_margin = Cm(2.0)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_after = Pt(2)
    _apply_tnr(h1.add_run("T.C.\nSAĞLIK BAKANLIĞI\nSATIN ALMA BİRİMİ"), size_pt=12, bold=True)

    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.paragraph_format.space_before = Pt(6)
    h2.paragraph_format.space_after = Pt(12)
    _apply_tnr(h2.add_run("TEKLİFE DAVET MEKTUBU"), size_pt=12, bold=True)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    meta.paragraph_format.space_after = Pt(10)
    _apply_tnr(meta.add_run("Konu: "), size_pt=12, bold=True)
    konu_text = f"4734/22-d Doğrudan Temin Piyasa Fiyat Araştırması ({procurement_type} Alımı)"
    if case_title.strip():
        konu_text = f"{case_title.strip()} ({procurement_type} Alımı)"
    _apply_tnr(meta.add_run(konu_text), size_pt=12, bold=False)

    p1_text, p2_text = build_invitation_text(deadline_date, deadline_time)
    body_p1 = doc.add_paragraph()
    body_p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body_p1.paragraph_format.line_spacing = 1.15
    body_p1.paragraph_format.first_line_indent = Cm(1.0)
    body_p1.paragraph_format.space_after = Pt(6)
    _apply_tnr(body_p1.add_run(p1_text), size_pt=12, bold=False)

    body_p2 = doc.add_paragraph()
    body_p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body_p2.paragraph_format.first_line_indent = Cm(1.0)
    body_p2.paragraph_format.space_after = Pt(14)
    _apply_tnr(body_p2.add_run(p2_text), size_pt=12, bold=False)

    second_col = "Hizmetin Adı ve Açıklaması" if procurement_type == "Hizmet" else "Malzeme Kaleminin Adı"
    headers = ["Sıra No", second_col, "Birim", "Miktar", "Birim Fiyat", "Toplam Fiyat"]
    col_widths = [Cm(1.3), Cm(7.2), Cm(1.8), Cm(1.8), Cm(2.4), Cm(2.5)]

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table)

    hdr_row = table.rows[0]
    hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    for i, h_title in enumerate(headers):
        _format_cell(hdr_row.cells[i], h_title, col_widths[i], WD_ALIGN_PARAGRAPH.CENTER, size_pt=11, bold=True, shade="E2E8F0")

    for idx, row_data in enumerate(table_rows, 1):
        row = table.add_row()
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        item_val = (
            row_data.get(second_col)
            or row_data.get("Malzeme Kaleminin Adı")
            or row_data.get("Hizmetin Adı ve Açıklaması")
            or row_data.get("name")
            or ""
        )
        raw_qty = row_data.get("Miktar") if "Miktar" in row_data else row_data.get("quantity", "")
        qty_str = format_qty(raw_qty) if isinstance(raw_qty, (int, float)) else str(raw_qty or "")
        vals = [
            str(row_data.get("Sıra No") or idx),
            str(item_val),
            str(row_data.get("Birim") or row_data.get("unit") or ""),
            qty_str,
            str(row_data.get("Birim Fiyat") or ""),
            str(row_data.get("Toplam Fiyat") or ""),
        ]
        aligns = [
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ]
        for i, val in enumerate(vals):
            _format_cell(row.cells[i], val, col_widths[i], aligns[i], size_pt=11, bold=False)

    total_row = table.add_row()
    total_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    merged_cell = total_row.cells[0].merge(total_row.cells[4])
    _format_cell(merged_cell, "GENEL TOPLAM (KDV Hariç)", sum(col_widths[:5], Cm(0)), WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11, bold=True, shade="F8FAFC")
    _format_cell(total_row.cells[5], "", col_widths[5], WD_ALIGN_PARAGRAPH.RIGHT, size_pt=11, bold=True, shade="F8FAFC")

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(18)
    spacer.paragraph_format.space_after = Pt(4)

    sign_p = doc.add_paragraph()
    sign_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sign_p.paragraph_format.space_before = Pt(12)
    _apply_tnr(sign_p.add_run("Teklif Veren Firma Kaşe / Yetkili İmza\n\n"), size_pt=12, bold=True)
    _apply_tnr(sign_p.add_run("Adı Soyadı / Ünvanı : ...........................................\n"), size_pt=12, bold=False)
    _apply_tnr(sign_p.add_run("Tarih               : .... / .... / 2026"), size_pt=12, bold=False)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Adım 4: Modüler Piyasa Fiyat Araştırma Tutanağı (.docx)
# ---------------------------------------------------------------------------

def prepare_market_research_context(
    case_data: dict[str, Any],
    offers: list[dict[str, Any]],
    evaluation: dict[str, Any],
    commission_members: list[dict[str, str]],
    authorizer_name: str,
    authorizer_title: str,
    report_date: str = "",
) -> dict[str, Any]:
    """Özel şablon (.docx) veya standart tutanak üreticisi için ortak veri bağlamı hazırlar."""
    date_str = report_date.strip() or dt.date.today().strftime("%d.%m.%Y")
    awarded_lines = [
        f"- {a['firm_name']}: {a['item_count']} kalem · Toplam {format_tr_number(a['total_amount'])} TL"
        for a in evaluation.get("awarded_summary", [])
    ]
    return {
        "ALIM_ADI": case_data.get("title") or "Satın Alma Dosyası",
        "ALIM_TURU": case_data.get("procurement_type") or "Malzeme",
        "DEGERLENDIRME_YONTEMI": evaluation.get("evaluation_mode") or "Kalem Bazlı En Düşük",
        "TARIH": date_str,
        "GENEL_UHDE_TOPLAMI": f"{format_tr_number(evaluation.get('grand_awarded_total', 0.0))} TL",
        "UHDE_OZETI": "\n".join(awarded_lines) if awarded_lines else "Henüz uygun teklif belirlenmedi.",
        "HARCAMA_YETKILISI_AD": authorizer_name.strip() or "........................................",
        "HARCAMA_YETKILISI_UNVAN": authorizer_title.strip() or "Başhekim / Harcama Yetkilisi",
        "KOMISYON_UYELERI": [
            {
                "name": (m.get("name") or "").strip() or "................................",
                "title": (m.get("title") or "").strip() or "Komisyon Üyesi",
            }
            for m in commission_members[:3]
        ],
        "offers": offers,
        "evaluation": evaluation,
    }


def _replace_placeholders_in_doc(doc: Any, mapping: dict[str, str]) -> None:
    """Özel şablon .docx dosyası bağlandığında metin/tablo yer tutucularını değiştirir."""
    def replace_in_paragraph(p: Any) -> None:
        full_text = p.text
        if not any(k in full_text for k in mapping):
            return
        new_text = full_text
        for k, v in mapping.items():
            new_text = new_text.replace(f"{{{{{k}}}}}", v)
        if p.runs:
            p.runs[0].text = new_text
            for r in p.runs[1:]:
                r.text = ""

    for p in doc.paragraphs:
        replace_in_paragraph(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)


def _append_comparison_table(doc: Any, ctx: dict[str, Any], usable_width_cm: float) -> None:
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Cm

    offers: list[dict[str, Any]] = ctx["offers"]
    evaluation: dict[str, Any] = ctx["evaluation"]
    item_results: list[dict[str, Any]] = evaluation.get("item_results", [])
    firm_totals: dict[int, dict[str, Any]] = evaluation.get("firm_totals", {})
    total_winner_id = evaluation.get("total_winner_id")

    n_firms = len(offers)
    # Sütunlar: Sıra, Kalem Adı, Birim, Miktar, [Firma 1..N], Uhde Firma, En Uygun Toplam
    headers = (
        ["Sıra", "Malzeme / Hizmet Adı", "Birim", "Miktar"]
        + [f"{o['firm_name']}\n(Birim / Toplam)" for o in offers]
        + ["Uhde (En Uygun Firma)", "En Uygun Toplam"]
    )

    fixed_w = 1.0 + 4.8 + 1.4 + 1.4 + 3.2 + 2.4
    rem_w = max(usable_width_cm - fixed_w, 3.0 * max(n_firms, 1))
    firm_w = round(rem_w / max(n_firms, 1), 2) if n_firms else 0.0
    col_widths_cm = [1.0, 4.8, 1.4, 1.4] + [firm_w] * n_firms + [3.2, 2.4]
    scale = usable_width_cm / sum(col_widths_cm)
    col_widths = [Cm(round(w * scale, 2)) for w in col_widths_cm]

    font_sz = 9 if n_firms >= 4 else 10
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table, color="CBD5E1", sz="6")

    hdr_row = table.rows[0]
    hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    for i, h_title in enumerate(headers):
        _format_cell(hdr_row.cells[i], h_title, col_widths[i], WD_ALIGN_PARAGRAPH.CENTER, size_pt=font_sz, bold=True, shade="E2E8F0")

    for res in item_results:
        it = res["item"]
        row = table.add_row()
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        winner_oid = res.get("winner_offer_id")

        _format_cell(row.cells[0], str(it["sira_no"]), col_widths[0], WD_ALIGN_PARAGRAPH.CENTER, size_pt=font_sz)
        _format_cell(row.cells[1], str(it["name"]), col_widths[1], WD_ALIGN_PARAGRAPH.LEFT, size_pt=font_sz)
        _format_cell(row.cells[2], str(it["unit"]), col_widths[2], WD_ALIGN_PARAGRAPH.CENTER, size_pt=font_sz)
        _format_cell(row.cells[3], format_qty(it["quantity"]), col_widths[3], WD_ALIGN_PARAGRAPH.CENTER, size_pt=font_sz)

        for idx_f, off in enumerate(offers):
            oid = int(off["id"])
            fp = res["firm_prices"].get(oid, {})
            up = fp.get("unit_price")
            lt = fp.get("line_total")
            is_win = winner_oid is not None and oid == int(winner_oid)
            txt = f"{format_tr_number(up)} TL\n({format_tr_number(lt)} TL)" if up else "-"
            _format_cell(
                row.cells[4 + idx_f],
                txt,
                col_widths[4 + idx_f],
                WD_ALIGN_PARAGRAPH.RIGHT,
                size_pt=font_sz,
                bold=is_win,
                shade="DCFCE7" if is_win else "",
                color_rgb=(22, 101, 52) if is_win else (0, 0, 0),
            )

        uhde_idx = 4 + n_firms
        _format_cell(
            row.cells[uhde_idx],
            res.get("winner_firm") or "-",
            col_widths[uhde_idx],
            WD_ALIGN_PARAGRAPH.LEFT,
            size_pt=font_sz,
            bold=bool(res.get("winner_firm")),
            shade="F0FDF4" if res.get("winner_firm") else "",
            color_rgb=(15, 118, 110) if res.get("winner_firm") else (0, 0, 0),
        )
        win_tot = res.get("winner_total_price")
        _format_cell(
            row.cells[uhde_idx + 1],
            f"{format_tr_number(win_tot)} TL" if win_tot else "-",
            col_widths[uhde_idx + 1],
            WD_ALIGN_PARAGRAPH.RIGHT,
            size_pt=font_sz,
            bold=bool(win_tot),
            shade="F0FDF4" if win_tot else "",
        )

    # Genel Toplam Satırı
    tot_row = table.add_row()
    tot_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    merged_lbl = tot_row.cells[0].merge(tot_row.cells[3])
    _format_cell(
        merged_lbl,
        "GENEL TOPLAM (KDV Hariç)",
        sum(col_widths[:4], Cm(0)),
        WD_ALIGN_PARAGRAPH.RIGHT,
        size_pt=font_sz,
        bold=True,
        shade="F1F5F9",
    )
    for idx_f, off in enumerate(offers):
        oid = int(off["id"])
        ft = firm_totals.get(oid, {})
        is_tot_win = total_winner_id is not None and oid == int(total_winner_id)
        _format_cell(
            tot_row.cells[4 + idx_f],
            f"{format_tr_number(ft.get('grand_total', 0.0))} TL",
            col_widths[4 + idx_f],
            WD_ALIGN_PARAGRAPH.RIGHT,
            size_pt=font_sz,
            bold=True,
            shade="DCFCE7" if is_tot_win else "F1F5F9",
            color_rgb=(22, 101, 52) if is_tot_win else (0, 0, 0),
        )
    uhde_idx = 4 + n_firms
    _format_cell(
        tot_row.cells[uhde_idx],
        "UHDE TOPLAMI",
        col_widths[uhde_idx],
        WD_ALIGN_PARAGRAPH.RIGHT,
        size_pt=font_sz,
        bold=True,
        shade="ECFDF5",
        color_rgb=(15, 118, 110),
    )
    _format_cell(
        tot_row.cells[uhde_idx + 1],
        ctx["GENEL_UHDE_TOPLAMI"],
        col_widths[uhde_idx + 1],
        WD_ALIGN_PARAGRAPH.RIGHT,
        size_pt=font_sz,
        bold=True,
        shade="ECFDF5",
        color_rgb=(15, 118, 110),
    )


def _append_signatures(doc: Any, ctx: dict[str, Any], usable_width_cm: float) -> None:
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.shared import Cm, Pt

    members: list[dict[str, str]] = ctx["KOMISYON_UYELERI"]
    if not members:
        members = [{"name": "................................", "title": "Komisyon Üyesi"}]

    sec_title = doc.add_paragraph()
    sec_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sec_title.paragraph_format.space_before = Pt(14)
    sec_title.paragraph_format.space_after = Pt(8)
    _apply_tnr(sec_title.add_run("PİYASA FİYAT ARAŞTIRMA KOMİSYONU"), size_pt=11, bold=True)

    n_mem = len(members)
    sig_table = doc.add_table(rows=1, cols=n_mem)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_table.autofit = False
    _remove_table_borders(sig_table)
    col_w = Cm(round(usable_width_cm / n_mem, 2))
    row = sig_table.rows[0]
    row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for i, mem in enumerate(members):
        cell = row.cells[i]
        cell.width = col_w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        _apply_tnr(p.add_run(f"Üye {i + 1}\n"), size_pt=10, bold=True)
        _apply_tnr(p.add_run(f"{mem['name']}\n"), size_pt=11, bold=True)
        _apply_tnr(p.add_run(f"{mem['title']}\n\n"), size_pt=10, bold=False)
        _apply_tnr(p.add_run("İmza: .................."), size_pt=10, bold=False)

    # Alt kısımda Hastane Harcama Yetkilisi (Başhekim)
    auth_p = doc.add_paragraph()
    auth_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    auth_p.paragraph_format.space_before = Pt(18)
    auth_p.paragraph_format.space_after = Pt(2)
    _apply_tnr(auth_p.add_run("UYGUNDUR\nHarcama Yetkilisi\n\n"), size_pt=11, bold=True)
    _apply_tnr(auth_p.add_run(f"{ctx['HARCAMA_YETKILISI_AD']}\n"), size_pt=11, bold=True)
    _apply_tnr(auth_p.add_run(f"{ctx['HARCAMA_YETKILISI_UNVAN']}"), size_pt=10, bold=False)


def build_market_research_docx(
    case_data: dict[str, Any],
    offers: list[dict[str, Any]],
    evaluation: dict[str, Any],
    commission_members: list[dict[str, str]],
    authorizer_name: str,
    authorizer_title: str,
    report_date: str = "",
    template_path: Path | None = None,
) -> bytes:
    """Piyasa Fiyat Araştırma Tutanağı (.docx) üretir.
    İleride özel kurum şablonu bağlanmak istenirse `template_path` parametresi üzerinden çalışır.
    """
    import docx
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt

    ctx = prepare_market_research_context(
        case_data=case_data,
        offers=offers,
        evaluation=evaluation,
        commission_members=commission_members,
        authorizer_name=authorizer_name,
        authorizer_title=authorizer_title,
        report_date=report_date,
    )

    if template_path is not None and template_path.exists():
        doc = docx.Document(str(template_path))
        mapping = {k: str(v) for k, v in ctx.items() if isinstance(v, (str, int, float))}
        _replace_placeholders_in_doc(doc, mapping)
        _append_comparison_table(doc, ctx, usable_width_cm=17.0)
        _append_signatures(doc, ctx, usable_width_cm=17.0)
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    doc = docx.Document()
    usable_width_cm = 25.7  # A4 Yatay (29.7 - 2 - 2)
    for sec in doc.sections:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width = Cm(29.7)
        sec.page_height = Cm(21.0)
        sec.top_margin = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin = Cm(2.0)
        sec.right_margin = Cm(2.0)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(11)

    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_after = Pt(4)
    _apply_tnr(h1.add_run("T.C. SAĞLIK BAKANLIĞI\nPİYASA FİYAT ARAŞTIRMA TUTANAĞI"), size_pt=13, bold=True)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    meta.paragraph_format.space_before = Pt(6)
    meta.paragraph_format.space_after = Pt(10)
    _apply_tnr(meta.add_run("Alım Adı: "), size_pt=11, bold=True)
    _apply_tnr(meta.add_run(f"{ctx['ALIM_ADI']}   |   "), size_pt=11)
    _apply_tnr(meta.add_run("Alım Türü: "), size_pt=11, bold=True)
    _apply_tnr(meta.add_run(f"{ctx['ALIM_TURU']}   |   "), size_pt=11)
    _apply_tnr(meta.add_run("Değerlendirme Usulü: "), size_pt=11, bold=True)
    _apply_tnr(meta.add_run(f"{ctx['DEGERLENDIRME_YONTEMI']}   |   "), size_pt=11)
    _apply_tnr(meta.add_run("Tarih: "), size_pt=11, bold=True)
    _apply_tnr(meta.add_run(ctx["TARIH"]), size_pt=11)

    _append_comparison_table(doc, ctx, usable_width_cm=usable_width_cm)

    dec_p = doc.add_paragraph()
    dec_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    dec_p.paragraph_format.space_before = Pt(10)
    dec_p.paragraph_format.space_after = Pt(4)
    _apply_tnr(dec_p.add_run("UHDE KARARI VE ÖZETİ:\n"), size_pt=11, bold=True)
    _apply_tnr(
        dec_p.add_run(
            f"4734 sayılı Kamu İhale Kanunu'nun 22/d maddesi kapsamında yapılan piyasa fiyat araştırması sonucunda "
            f"({ctx['DEGERLENDIRME_YONTEMI']} esas alınarak) toplam {ctx['GENEL_UHDE_TOPLAMI']} (KDV Hariç) bedelle "
            f"aşağıdaki şekilde alım yapılmasına karar verilmiştir:\n{ctx['UHDE_OZETI']}"
        ),
        size_pt=10,
    )

    _append_signatures(doc, ctx, usable_width_cm=usable_width_cm)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

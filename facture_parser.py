"""
facture_parser.py
------------------
Specialized parser for the "FACTURE PROFORMA" invoice template used by
LAHLAH KARIM (Transport de marchandise publiques).

Two source-document variants have been seen in the wild:
  - A clean, plain-text PDF (early invoices).
  - A PDF printed from an HTML page (later invoices, "LAHLAHV3" template).
    In this variant, when the "Désignation" text is long, the Qté / P.U HT
    spinner-widget values are rendered overlapping the description text,
    which corrupts a flat text extraction of that line. The table grid
    (extracted via pdfplumber's table detector) is much more reliable
    here: the Date and Montant HT *columns* stay clean even when the
    Désignation/Qté/P.U HT columns in between get garbled — so line items
    are parsed from the table structure, not the flat text.

If a PDF does not contain "FACTURE PROFORMA", parse_facture() returns
None and the generic extractor (extractor.py) is used as a fallback,
so the app still works on other PDF types.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class LineItem:
    date: str
    designation: str
    qte: float
    pu_ht: float
    montant_ht: float


@dataclass
class Facture:
    file_name: str = ""
    societe: str = ""
    societe_activite: str = ""
    societe_adresse: str = ""
    societe_tel: str = ""
    societe_rc: str = ""
    societe_nif: str = ""
    societe_article_imposition: str = ""
    societe_banque: str = ""
    numero: str = ""
    date: str = ""
    demandeur: str = ""
    client_nom: str = ""
    client_nif: str = ""
    client_adresse: str = ""
    lignes: list[LineItem] = field(default_factory=list)
    sous_total_ht: float | None = None
    tva_percent: float | None = None
    tva_montant: float | None = None
    total_ttc: float | None = None
    devise: str = "DZD"


AMOUNT_RE = re.compile(r"([\d\s\u202f\xa0\ufffd]+,\d{2})\s*([A-Z]{2,4})?")
DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")


def _to_float(amount_str: str) -> float:
    """
    '23 000,00' -> 23000.0 (handles thin/normal spaces as thousands separators).

    Also recovers a specific, reproducible font-decoding glitch seen in the
    "LAHLAHV3" PDF template: a leading "1" digit sometimes extracts as the
    Unicode replacement character "\ufffd" (e.g. "19 500,00" -> "�9 500,00",
    "12 000,00" -> "�2 000,00"). Since this only ever happens as the FIRST
    character of the number, replacing a leading "\ufffd" with "1" is safe.
    """
    cleaned = amount_str.replace("\u202f", " ").replace("\xa0", " ").strip()
    if cleaned.startswith("\ufffd"):
        cleaned = "1" + cleaned[1:]
    cleaned = re.sub(r"\s", "", cleaned)
    cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def _clean_address(raw: str) -> str:
    # Strip private-use-area glyphs (the little scroll arrows in the PDF widget) and blank lines
    cleaned = re.sub(r"[\ue000-\uf8ff]", "", raw)
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]
    return ", ".join(lines)


def _clean_designation(raw: str) -> str:
    """Collapse the joined designation/qté/p.u-ht cell text into one tidy string."""
    cleaned = re.sub(r"\s+", " ", raw).strip()
    cleaned = cleaned.strip(" -")
    return cleaned


def is_facture_proforma(text: str) -> bool:
    return "FACTURE PROFORMA" in text.upper()


def _find_designation_table(tables):
    """Find the "Date / Désignation / Qté / P.U HT / Montant HT" table among a PdfResult's tables."""
    for t in tables or []:
        headers = [h.strip() for h in t.headers if h]
        if len(headers) < 4:
            continue
        if headers[0].lower() != "date":
            continue
        joined = " ".join(headers).lower()
        if ("désignation" in joined or "designation" in joined) and "montant" in joined:
            return t
    return None


def _parse_line_items_from_table(table) -> list[LineItem]:
    lignes: list[LineItem] = []
    for row in table.rows:
        cells = [c.strip() if c else "" for c in row]
        if not cells or not cells[0]:
            continue
        date_cell = cells[0]
        if not DATE_RE.match(date_cell):
            # Skip the unused placeholder row ("mm/dd/yyyy ...") and any
            # stray fragment rows the table detector picked up outside the grid.
            continue

        # Montant HT is the last non-empty cell that looks like a money amount.
        montant_ht = None
        montant_idx = None
        for idx in range(len(cells) - 1, 0, -1):
            m = AMOUNT_RE.search(cells[idx])
            if m:
                try:
                    montant_ht = _to_float(m.group(1))
                    montant_idx = idx
                    break
                except ValueError:
                    continue
        if montant_ht is None:
            continue

        # Everything between the date and the montant column is the
        # désignation + (possibly garbled) qté/p.u-ht text — join it all as
        # a best-effort description, since qté/p.u-ht aren't used downstream.
        middle = " ".join(c for c in cells[1:montant_idx] if c)
        designation = _clean_designation(middle)
        if designation.lower() == "description de la prestation":
            continue  # unused template placeholder row

        lignes.append(LineItem(date=date_cell, designation=designation, qte=1.0, pu_ht=montant_ht, montant_ht=montant_ht))
    return lignes


def parse_facture(text: str, file_name: str = "", tables=None) -> Facture | None:
    if not is_facture_proforma(text):
        return None

    f = Facture(file_name=file_name)

    # ---- Company header ---------------------------------------------------
    m = re.search(r"FACTURE PROFORMA\s*\n([A-Za-zÀ-ÿ' \-]+?)\s*N°", text)
    if m:
        f.societe = m.group(1).strip()

    m = re.search(r"N°\s*\n\s*(\S+)", text)
    if m:
        f.numero = m.group(1).strip()

    m = re.search(r"N°\s*\n\s*\S+\s*\n(.+?)\n", text)
    if m:
        f.societe_activite = m.group(1).strip()

    # The invoice date always sits on its own line right before "Tél" —
    # true in both the older template (Date label, then address, then date
    # value) and the newer "LAHLAHV3" template (address + "Date" label on
    # the same line, then the date value on the next line).
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})\s*\nTél", text)
    if m:
        f.date = m.group(1).strip()

    m = re.search(r"Date\n(.+?)\n\d{1,2}/\d{1,2}/\d{2,4}\nTél", text)
    if m:
        f.societe_adresse = m.group(1).strip()

    m = re.search(r"Tél\s*:\s*([^\n]+)", text)
    if m:
        f.societe_tel = m.group(1).strip()

    m = re.search(r"RC\s*:\s*(.+?)(?:\s+Demandeur)?\n", text)
    if m:
        f.societe_rc = m.group(1).strip()

    m = re.search(r"Demandeur\s*\n([^\n]+)", text)
    if m:
        f.demandeur = m.group(1).strip()

    # Company NIF has a colon; client NIF (below) does not, so this only matches the company one.
    m = re.search(r"\bNIF\s*:\s*(\S+)", text)
    if m:
        f.societe_nif = m.group(1).strip()

    m = re.search(r"Article d'imposition\s*:\s*(\S+)", text)
    if m:
        f.societe_article_imposition = m.group(1).strip()

    m = re.search(r"Banque\s*:\s*([^\n]+)", text)
    if m:
        f.societe_banque = m.group(1).strip()

    # ---- Client -------------------------------------------------------------
    m = re.search(r"Nom NIF\s*\n(.+?)\s+(\d{6,})\s*\n", text)
    if m:
        f.client_nom = m.group(1).strip()
        f.client_nif = m.group(2).strip()

    m = re.search(r"Adresse\s*\n(.+?)\n(?:DÉSIGNATION|DESIGNATION)", text, re.S)
    if m:
        f.client_adresse = _clean_address(m.group(1))

    # ---- Line items ---------------------------------------------------------
    # Prefer the table grid (robust against overlapping/garbled cell text);
    # fall back to a flat-text line regex only if no matching table was found.
    table = _find_designation_table(tables)
    if table is not None:
        f.lignes = _parse_line_items_from_table(table)

    if not f.lignes:
        m = re.search(
            r"Date Désignation Qté P\.U HT Montant HT\s*\n(.+?)\nSous-total HT",
            text,
            re.S,
        )
        if m:
            block = m.group(1)
            line_re = re.compile(
                r"(\d{1,2}/\d{1,2}/\d{2,4})\s+(.+?)\s+(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)\s+"
                r"([\d\s\u202f\xa0]+,\d{2})\s*([A-Z]{2,4})?\s*$"
            )
            for line in block.splitlines():
                line = line.strip()
                if not line:
                    continue
                lm = line_re.match(line)
                if lm:
                    date_, designation, qte, pu_ht, montant, devise = lm.groups()
                    if designation.strip().lower() == "description de la prestation":
                        continue
                    try:
                        f.lignes.append(
                            LineItem(
                                date=date_,
                                designation=designation.strip(),
                                qte=_to_float(qte),
                                pu_ht=_to_float(pu_ht),
                                montant_ht=_to_float(montant),
                            )
                        )
                        if devise:
                            f.devise = devise
                    except ValueError:
                        continue

    # ---- Totals -----------------------------------------------------------
    m = re.search(r"Sous-total HT\s+([\d\s\u202f\xa0\ufffd]+,\d{2})\s*([A-Z]{2,4})?", text)
    if m:
        f.sous_total_ht = _to_float(m.group(1))
        if m.group(2):
            f.devise = m.group(2)

    m = re.search(r"TVA\s+(\d+(?:[.,]\d+)?)\s*%\s+([\d\s\u202f\xa0\ufffd]+,\d{2})", text)
    if m:
        f.tva_percent = _to_float(m.group(1))
        f.tva_montant = _to_float(m.group(2))

    m = re.search(r"Total TTC\s+([\d\s\u202f\xa0\ufffd]+,\d{2})", text)
    if m:
        f.total_ttc = _to_float(m.group(1))

    return f

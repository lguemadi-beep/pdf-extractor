"""
facture_parser.py
------------------
Specialized parser for the "FACTURE PROFORMA" invoice template used by
LAHLAH KARIM (Transport de marchandise publiques). These are clean,
consistent, text-based PDFs, so we parse them precisely into structured
fields instead of relying on generic key-value guessing.

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


def _to_float(amount_str: str) -> float:
    """'23 000,00' -> 23000.0 (handles thin/normal spaces as thousands separators)."""
    cleaned = amount_str.replace("\u202f", " ").replace("\xa0", " ")
    cleaned = re.sub(r"\s", "", cleaned)
    cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def _clean_address(raw: str) -> str:
    # Strip private-use-area glyphs (the little scroll arrows in the PDF widget) and blank lines
    cleaned = re.sub(r"[\ue000-\uf8ff]", "", raw)
    lines = [ln.strip() for ln in cleaned.splitlines() if ln.strip()]
    return ", ".join(lines)


def is_facture_proforma(text: str) -> bool:
    return "FACTURE PROFORMA" in text.upper()


def parse_facture(text: str, file_name: str = "") -> Facture | None:
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

    m = re.search(r"N°\s*\n\s*\S+\s*\n(.+?)\nDate\n", text)
    if m:
        f.societe_activite = m.group(1).strip()

    # The company address line sits between "Date" and the actual date value
    # (the two-column layout interleaves label/value by vertical position).
    m = re.search(r"Date\n(.+?)\n(\d{1,2}/\d{1,2}/\d{2,4})\nTél", text)
    if m:
        f.societe_adresse = m.group(1).strip()
        f.date = m.group(2).strip()

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

    # ---- Line items -----------------------------------------------------
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
    m = re.search(r"Sous-total HT\s+([\d\s\u202f\xa0]+,\d{2})\s*([A-Z]{2,4})?", text)
    if m:
        f.sous_total_ht = _to_float(m.group(1))
        if m.group(2):
            f.devise = m.group(2)

    m = re.search(r"TVA\s+(\d+(?:[.,]\d+)?)\s*%\s+([\d\s\u202f\xa0]+,\d{2})", text)
    if m:
        f.tva_percent = _to_float(m.group(1))
        f.tva_montant = _to_float(m.group(2))

    m = re.search(r"Total TTC\s+([\d\s\u202f\xa0]+,\d{2})", text)
    if m:
        f.total_ttc = _to_float(m.group(1))

    return f
